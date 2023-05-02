' Skeleton Program for the AQA AS Summer 2023 examination
' this code should be used in conjunction with the Preliminary Material
' written by the AQA Programmer Team
' developed in the Visual Studio Community Edition environment

' Version number: 0.0.0

Imports System.IO

Module Module1
    Const EMPTY_STRING As String = ""
    Const HI_MEM As Integer = 20
    Const MAX_INT As Integer = 127 ' 8 bits available for operand (two's complement integer)
    Const PC As Integer = 0
    Const ACC As Integer = 1
    Const STATUS As Integer = 2
    Const TOS As Integer = 3
    Const ERR As Integer = 4

    ''' <summary>
    ''' class data structure to hold OpCode, OperandString, and OperandValue
    ''' </summary>
    Structure AssemblerInstruction
        Dim OpCode As String
        Dim OperandString As String
        Dim OperandValue As Integer
    End Structure

    ''' <summary>
    ''' displays menu options
    ''' </summary>
    Sub DisplayMenu()
        Console.WriteLine()
        Console.WriteLine("Main Menu")
        Console.WriteLine("=========")
        Console.WriteLine("L - Load a program file")
        Console.WriteLine("D - Display source code")
        Console.WriteLine("E - Edit source code")
        Console.WriteLine("A - Assemble program")
        Console.WriteLine("R - Run the program")
        Console.WriteLine("X - Exit simulator")
        Console.WriteLine()
    End Sub

    ''' <summary>
    ''' grabs and returns user input char
    ''' </summary>
    ''' <returns></returns>
    Function GetMenuOption() As Char
        Dim Choice As String = EMPTY_STRING
        While Choice.Length <> 1
            Console.Write("Enter your choice: ")
            Choice = Console.ReadLine().ToUpper()    ' added to upper
        End While
        Return Choice(0)
    End Function

    ''' <summary>
    ''' initialises sourceCode variable to hold 20 empty strings
    ''' </summary>
    ''' <param name="SourceCode"></param>
    Sub ResetSourceCode(SourceCode() As String)
        For LineNumber = 0 To HI_MEM - 1
            SourceCode(LineNumber) = EMPTY_STRING
        Next
    End Sub

    ''' <summary>
    ''' initialises memory variable to have null values
    ''' </summary>
    ''' <param name="Memory"></param>
    Sub ResetMemory(Memory() As AssemblerInstruction)
        For LineNumber = 0 To HI_MEM - 1
            Memory(LineNumber).OpCode = EMPTY_STRING
            Memory(LineNumber).OperandString = EMPTY_STRING
            Memory(LineNumber).OperandValue = 0
        Next
    End Sub
    ''' <summary>
    ''' displays assembly source coded loaded from textfile. 
    ''' </summary>
    ''' <param name="SourceCode"></param>
    Sub DisplaySourceCode(SourceCode() As String)
        Console.WriteLine()
        Dim NumberOfLines As Integer = Convert.ToInt32(SourceCode(0))
        For LineNumber = 0 To NumberOfLines
            Console.WriteLine($"{LineNumber,2} {SourceCode(LineNumber),-40}")       ' displays SourceCode line by line for all lines
        Next
        Console.WriteLine()
    End Sub
    ''' <summary>
    ''' asks for filename and then reads the file line by line. Saves it in SourceCode()
    ''' </summary>
    ''' <param name="SourceCode"></param>
    Sub LoadFile(SourceCode() As String)
        Dim FileExists As Boolean = False
        ResetSourceCode(SourceCode)            ' incase there is residual date from previous execution
        Dim LineNumber As Integer = 0
        Console.Write("Enter filename to load: ")
        Dim FileName As String = Console.ReadLine()
        Try
            Dim MyReader As StreamReader = New StreamReader(FileName & ".txt")    ' if no file, throw fatal error (do the catch part)
            FileExists = True                                                    ' if previous fail no fail, means file is found so fileExists = true
            Dim Instruction As String = MyReader.ReadLine()
            While Not Instruction Is Nothing  ' while instructions is not empty
                LineNumber += 1
                SourceCode(LineNumber) = Instruction  ' loads each line from file into SourceCode
                Instruction = MyReader.ReadLine()      ' read next line 
            End While
            MyReader.Close()   ' close when done
            SourceCode(0) = Convert.ToString(LineNumber)   ' first line is number of lines of code in assmbly file
        Catch
            If Not FileExists Then
                Console.WriteLine("Error Code 1")      ' no file, no file reading
            Else
                Console.WriteLine("Error Code 2")          ' file? no code? no workie
                SourceCode(0) = Str(LineNumber - 1)
            End If
        End Try
        If LineNumber > 0 Then  ' if there like actual lines, display da code
            DisplaySourceCode(SourceCode)
        End If
    End Sub
    ''' <summary>
    ''' lets you edit source code line by line.
    ''' </summary>
    ''' <param name="SourceCode"></param>
    Sub EditSourceCode(SourceCode() As String)
        Dim Choice As String = EMPTY_STRING
        Dim LineNumber As Integer
        Console.Write("Enter line number of code to edit: ")
        LineNumber = Console.ReadLine()
        Console.WriteLine(SourceCode(LineNumber))  ' you choose which line to edit and it prints out
        While Choice <> "C"  ' while not C, therefore c is cancel out loop
            Choice = EMPTY_STRING      ' reset string on each run
            While Choice <> "E" And Choice <> "C"     ' guard against miss input
                Console.WriteLine("E - Edit this line")
                Console.WriteLine("C - Cancel edit")
                Console.Write("Enter your choice: ")
                Choice = Console.ReadLine()
            End While
            If Choice = "E" Then
                Console.Write("Enter the new line: ")
                SourceCode(LineNumber) = Console.ReadLine()    ' changes code at that line
            End If
            DisplaySourceCode(SourceCode)
        End While
    End Sub

    ''' <summary>
    ''' Updates the symboltable dictionary which holds all the labels and lines where they occur
    ''' </summary>
    ''' <param name="SymbolTable"></param>
    ''' <param name="ThisLabel"></param>
    ''' <param name="LineNumber"></param>
    Sub UpdateSymbolTable(SymbolTable As Dictionary(Of String, Integer), ThisLabel As String, LineNumber As Integer)
        If SymbolTable.ContainsKey(ThisLabel) Then
            Console.WriteLine("Error Code 3")  ' if table has duplicate key, same name for different labels used - assembly code must have an error
        Else
            SymbolTable.Add(ThisLabel, LineNumber)
        End If
    End Sub

    ''' <summary>
    ''' extract label from the line of code inputted.-------
    ''' i.e ->  SUB1 in SUB1: ADD  NUM1. also saves line number where label is
    ''' </summary>
    ''' <param name="Instruction"></param>
    ''' <param name="LineNumber"></param>
    ''' <param name="Memory"></param>
    ''' <param name="SymbolTable"></param>
    Sub ExtractLabel(Instruction As String, LineNumber As Integer, Memory() As AssemblerInstruction, SymbolTable As Dictionary(Of String, Integer))
        If Instruction.Length > 0 Then ' if line not empty
            Dim ThisLabel As String = Instruction.Substring(0, 5).TrimStart(" ") 'this exacts the label itself
            If ThisLabel <> EMPTY_STRING Then '
                If Instruction(5) <> ":" Then ' error when : is in the label
                    Console.WriteLine("Error Code 4") ' error if label is too short.
                    Memory(0).OpCode = "ERR"
                Else
                    UpdateSymbolTable(SymbolTable, ThisLabel, LineNumber)
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' extract OpCode from the line of code inputted. This is the instruction-------
    ''' i.e ->  ADD in SUB1: ADD  NUM1
    ''' </summary>
    ''' <param name="Instruction"></param>
    ''' <param name="LineNumber"></param>
    ''' <param name="Memory"></param>
    Sub ExtractOpCode(Instruction As String, LineNumber As Integer, Memory() As AssemblerInstruction)
        If Instruction.Length > 9 Then ' if the instruction is a proper command
            Dim OpCodeValues() As String = {"LDA", "STA", "LDA#", "HLT", "ADD", "JMP", "SUB", "CMP#", "BEQ", "SKP", "JSR", "RTN", "   "}
            Dim Operation As String = Instruction.Substring(7, 3) ' extracts string from index 7 for 3 chars.
            If Instruction.Length > 10 Then                       ' if smth like  NUM1:      2 then the opcode is just "   "
                Dim AddressMode As String = Instruction(10) ' 10th index of a line will always be either
                If AddressMode = "#" Then                   ' " " or "#"
                    Operation += AddressMode 'if addressing mode is direct, add # onto the opcode
                End If
            End If
            If OpCodeValues.Contains(Operation) Then   'if the opcode is a value that exists in the codex (known command)
                Memory(LineNumber).OpCode = Operation 'add the opcode at the line number it was in the memory var
            Else
                If Operation <> EMPTY_STRING Then   ' if opcode is empty, error
                    Console.WriteLine("Error Code 5")
                    Memory(0).OpCode = "ERR"
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' extract Operand from the line of code inputted. This is the value.-------
    ''' i.e ->  NUM1 in SUB1: ADD  NUM1
    ''' </summary>
    ''' <param name="Instruction"></param>
    ''' <param name="LineNumber"></param>
    ''' <param name="Memory"></param>
    Sub ExtractOperand(Instruction As String, LineNumber As Integer, Memory() As AssemblerInstruction)
        If Instruction.Length >= 13 Then ' if lone contains operand
            Dim Operand As String = Instruction.Substring(12) ' grabs operand
            Dim ThisPosition As Integer = -1 ' any value bellow 0
            For Position = 0 To Operand.Length - 1 ' looks for comment
                If Operand(Position) = "*" Then
                    ThisPosition = Position   ' this position is position of comment
                End If
            Next
            If ThisPosition >= 0 Then
                Operand = Operand.Substring(0, ThisPosition - 1) ' remove the comment from the command
            End If
            Operand = Operand.Trim(" ") ' trim empty space
            Memory(LineNumber).OperandString = Operand  ' saves operandstring into memory
        End If
    End Sub

    ''' <summary>
    ''' loops through the entire program reading and extracting all the labels, operations (opcode), and data (operand).
    ''' </summary>
    ''' <param name="SourceCode"></param>
    ''' <param name="Memory"></param>
    ''' <param name="SymbolTable">Holds labels and the int value is the line number.</param>
    Sub PassOne(SourceCode() As String, Memory() As AssemblerInstruction, SymbolTable As Dictionary(Of String, Integer))
        Dim NumberOfLines As Integer = Convert.ToInt32(SourceCode(0))
        For LineNumber = 1 To NumberOfLines
            Dim Instruction As String = SourceCode(LineNumber)  ' instructions holds each line in the loop
            ExtractLabel(Instruction, LineNumber, Memory, SymbolTable)
            ExtractOpCode(Instruction, LineNumber, Memory)
            ExtractOperand(Instruction, LineNumber, Memory)
        Next
    End Sub

    ''' <summary>
    ''' deals with operandvalue. Asigns each operand string their value.
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="SymbolTable"></param>
    ''' <param name="NumberOfLines"></param>
    Sub PassTwo(Memory() As AssemblerInstruction, SymbolTable As Dictionary(Of String, Integer), NumberOfLines As Integer)
        Dim OperandValue As Integer
        For LineNumber = 1 To NumberOfLines  ' loop for all the lines
            Dim Operand As String = Memory(LineNumber).OperandString  ' grab the operand name in that line
            If Operand <> EMPTY_STRING Then  'if its not an empty string / there is an operand
                If SymbolTable.ContainsKey(Operand) Then  ' if that is an recognised operand
                    OperandValue = SymbolTable(Operand)   ' operandvalie = the interger value in the dict
                    Memory(LineNumber).OperandValue = OperandValue ' asign our operand value
                Else     ' if the command isnt recognised, must be a "register"
                    Try
                        OperandValue = Convert.ToInt32(Operand) ' convert it into an int
                        Memory(LineNumber).OperandValue = OperandValue ' asign operand value to 7
                    Catch
                        Console.WriteLine("Error Code 6")  ' otherwise there must be an error with the assembly code.
                        Memory(0).OpCode = "ERR"
                    End Try
                End If
            End If
        Next
    End Sub

    ' i mean the names say it all
    Sub DisplayMemoryLocation(Memory() As AssemblerInstruction, Location As Integer)
        Console.Write($"*  {Memory(Location).OpCode,-5}{Memory(Location).OperandValue,-5} |")
    End Sub

    Sub DisplaySourceCodeLine(SourceCode() As String, Location As Integer)
        Console.WriteLine($" {Location,3}  |  {SourceCode(Location),-40}")
    End Sub

    ''' <summary>
    ''' beginning of the process to display the code at each frame.
    ''' </summary>
    ''' <param name="SourceCode"></param>
    ''' <param name="Memory"></param>
    Sub DisplayCode(SourceCode() As String, Memory() As AssemblerInstruction)
        Console.WriteLine("*  Memory     Location  Label  Op   Operand Comment") ' displays the legend (if i can call it that)
        Console.WriteLine("*  Contents                    Code")
        Dim NumberOfLines As Integer = Convert.ToInt32(SourceCode(0))  ' sourceCode(0) always has number of lines in the program
        DisplayMemoryLocation(Memory, 0)
        Console.WriteLine("   0  |")
        For Location = 1 To NumberOfLines
            DisplayMemoryLocation(Memory, Location)
            DisplaySourceCodeLine(SourceCode, Location)
        Next
    End Sub

    ''' <summary>
    ''' assembles the memory. Populates Opcode, Operandstring, and Operandvalue for each line of code.
    ''' </summary>
    ''' <param name="SourceCode"></param>
    ''' <param name="Memory"></param>
    Sub Assemble(SourceCode() As String, Memory() As AssemblerInstruction)
        ResetMemory(Memory)
        Dim NumberOfLines As Integer = Convert.ToInt32(SourceCode(0)) ' GRABS NUMBER OF LINES OF ASSM CODE
        Dim SymbolTable As New Dictionary(Of String, Integer)         ' dict that holds
        PassOne(SourceCode, Memory, SymbolTable)
        If Memory(0).OpCode <> "ERR" Then ' if first line is an error
            Memory(0).OpCode = "JMP"  ' skip that line
            If SymbolTable.ContainsKey("START") Then  ' if theres a start label
                Memory(0).OperandValue = SymbolTable("START") ' set the first position in memory to be start
            Else
                Memory(0).OperandValue = 1 ' otherwise, start from line 1.
            End If
            PassTwo(Memory, SymbolTable, NumberOfLines) ' now we need to get our operandvalues
        End If
    End Sub


    ''' <summary>
    ''' Overcomplicated function to convert Denary to Binary
    ''' </summary>
    ''' <param name="DecimalNumber"></param>
    ''' <returns></returns>
    Function ConvertToBinary(DecimalNumber As Integer) As String
        Dim BinaryString As String = EMPTY_STRING                 ' initialise binary string
        While DecimalNumber > 0
            Dim Remainder As Integer = DecimalNumber Mod 2        ' grab remainder of our denary number / 2
            Dim Bit As String = Convert.ToString(Remainder)       ' makes the remainder a string called bit
            BinaryString = Bit + BinaryString                     ' adds on the bit to the start of the binary strinf
            DecimalNumber = DecimalNumber \ 2                     ' actually performs the division and saves the result back into the denary number
        End While
        While Len(BinaryString) < 3                               ' previous part onky adds 1s so here we need to add 0s
            BinaryString = "0" + BinaryString                     ' until the binary string is 3 bits (characters) long
        End While
        Return BinaryString
    End Function

    ''' <summary>
    ''' Name says it.
    ''' </summary>
    ''' <param name="BinaryString"></param>
    ''' <returns></returns>
    Function ConvertToDecimal(BinaryString As String) As Integer
        Dim DecimalNumber As Integer = 0                          ' initialise int variable called decicalNumber
        For Each Bit In BinaryString
            Dim BitValue As Integer = Convert.ToInt32(Bit) - 48   ' convert bit to 32 bit int. So if 1, 32int of 1 is 49. 32int of 0 is 48. TL;DR make the string an int
            DecimalNumber = DecimalNumber * 2 + BitValue          ' Decimalnumber * 2 + value for all bits will give decimal number.
        Next
        Return DecimalNumber                                      ' if you really wanna see how this works, hand tracetable it with inputs 100, 010, 001. Should get 4, 2, and 1.
    End Function

    ''' <summary>
    ''' displays each frame number before run
    ''' </summary>
    ''' <param name="FrameNumber"></param>
    Sub DisplayFrameDelimiter(FrameNumber As Integer)
        If FrameNumber = -1 Then
            Console.WriteLine("***************************************************************")
        Else
            Console.WriteLine($"****** Frame {FrameNumber} ************************************************")
        End If
    End Sub


    ''' <summary>
    ''' Displays the contents of Memory, the SourceCode, and the values of the different Registers
    ''' </summary>
    ''' <param name="SourceCode"></param>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    Sub DisplayCurrentState(SourceCode() As String, Memory() As AssemblerInstruction, Registers() As Integer)
        Console.WriteLine("*")
        DisplayCode(SourceCode, Memory)        ' this displays sourcecode and memory first 
        Console.WriteLine("*")
        Console.WriteLine($"*  PC:  {Registers(PC)}  ACC:  {Registers(ACC)}  TOS:  {Registers(TOS)}")      ' then displays all the register's values
        Console.WriteLine("*  Status Register: ZNV")
        Console.WriteLine($"*                   {ConvertToBinary(Registers(STATUS))}")
        DisplayFrameDelimiter(-1)
    End Sub

    ''' <summary> 
    ''' Sets flags depending on the value,
    ''' If Val == 0, Z is set (100)
    ''' If Val less than 0, N is set (010)
    ''' If Val is not representible in 8 bit binary, V is set (001) 
    ''' If Val is greater than 0, Nothing is set (000)
    ''' -----> REMEMBER: ZNV
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="Registers"></param>
    Sub SetFlags(Value As Integer, Registers() As Integer)
        If Value = 0 Then                                          ' sets flag 
            Registers(STATUS) = ConvertToDecimal("100")     ' Z = 1
        ElseIf Value < 0 Then
            Registers(STATUS) = ConvertToDecimal("010")     ' N = 1
        ElseIf Value > MAX_INT Or Value < -(MAX_INT + 1) Then
            Registers(STATUS) = ConvertToDecimal("001")     ' V == 1
        Else
            Registers(STATUS) = ConvertToDecimal("000")     ' Z, N, and V == 0
        End If
    End Sub

    ''' <summary>
    ''' Display a runtime error. Only ever used on overflows.
    ''' </summary>
    ''' <param name="ErrorMessage"></param>
    ''' <param name="Registers"></param>
    Sub ReportRunTimeError(ErrorMessage As String, Registers() As Integer)
        Console.WriteLine($"Run time error: {ErrorMessage}")
        Registers(ERR) = 1                                     ' will cause program to stop execution if there is an error
    End Sub

    ''' <summary>
    ''' loads accumulator with the value at the address specified
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteLDA(Memory() As AssemblerInstruction, Registers() As Integer, Address As Integer)
        Registers(ACC) = Memory(Address).OperandValue   ' loads accumulator with the value at the address specified
        SetFlags(Registers(ACC), Registers)             ' check for overflow, negative, positive, or 0
    End Sub

    ''' <summary>
    ''' Loads value in specified address into accumulator
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteSTA(Memory() As AssemblerInstruction, Registers() As Integer, Address As Integer)
        Memory(Address).OperandValue = Registers(ACC)   ' Loads value in specified address into accumulator
    End Sub

    ''' <summary>
    ''' loads accumulator with the immediate value specified
    ''' </summary>
    ''' <param name="Registers"></param>
    ''' <param name="Operand"></param>
    Sub ExecuteLDAimm(Registers() As Integer, Operand As Integer)
        Registers(ACC) = Operand                        ' loads accumulator with the immediate value specified
        SetFlags(Registers(ACC), Registers)             ' check for overflow, negative, positive, or 0
    End Sub

    ''' <summary>
    ''' adds value at the address specified to the accumulator
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteADD(Memory() As AssemblerInstruction, Registers() As Integer, Address As Integer)
        Registers(ACC) += Memory(Address).OperandValue             ' adds value at the address specified to the accumulator
        SetFlags(Registers(ACC), Registers)                        ' check for overflow, negative, positive, or 0
        If Registers(STATUS) = ConvertToDecimal("001") Then        ' if V is set (status is 1 (001))
            ReportRunTimeError("Overflow", Registers)              ' report overflow and close execution
        End If
    End Sub

    ''' <summary>
    ''' takes away value at the address specified from the accumulator
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteSUB(Memory() As AssemblerInstruction, Registers() As Integer, Address As Integer)
        Registers(ACC) -= Memory(Address).OperandValue             ' takes away value at the address specified from the accumulator
        SetFlags(Registers(ACC), Registers)                        ' check for overflow, negative, positive, or 0
        If Registers(STATUS) = ConvertToDecimal("001") Then        ' if V is set (status is 1 (001))
            ReportRunTimeError("Overflow", Registers)              ' report overflow and close execution
        End If
    End Sub

    ''' <summary>
    ''' ACC == operand is 100... ACC less than Operand is 010... ACC greater than Operand is 000
    ''' </summary>
    ''' <param name="Registers"></param>
    ''' <param name="Operand"></param>
    Sub ExecuteCMPimm(Registers() As Integer, Operand As Integer)
        Dim Value As Integer = Registers(ACC) - Operand
        SetFlags(Value, Registers)
    End Sub

    ''' <summary>
    ''' If flag Z is 1 (STATUS in binary is 100), perform a jump to address in operand
    ''' </summary>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteBEQ(Registers() As Integer, Address As Integer)
        Dim StatusRegister As String = ConvertToBinary(Registers(STATUS))
        Dim FlagZ As Char = StatusRegister(0)                               ' first 2 lines grab status, turn to binary, and get only first char (Z flag)
        If FlagZ = "1" Then
            Registers(PC) = Address                                         ' If Z is 1, perform the JMP to address specified
        End If
    End Sub

    ''' <summary>
    ''' Takes the OperandValue given and treats it as an address. Sets PC to that address
    ''' so it starts executing from there onwards
    ''' </summary>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteJMP(Registers() As Integer, Address As Integer)
        Registers(PC) = Address
    End Sub

    ''' <summary>
    ''' Its a skip
    ''' </summary>
    Sub ExecuteSKP()
    End Sub

    Sub DisplayStack(Memory() As AssemblerInstruction, Registers() As Integer)
        Console.WriteLine("Stack contents:")
        Console.WriteLine(" ----")
        For Index = Registers(TOS) To HI_MEM - 1
            Console.WriteLine($"|{Memory(Index).OperandValue,3} |")
        Next
        Console.WriteLine(" ----")
    End Sub

    ''' <summary>
    ''' Takes address and sets pc to that value (a jump to that address). Stores old PC in Memory from top-down
    ''' i.e: stores in 19, stores in 18, 17, 16, etc
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    ''' <param name="Address"></param>
    Sub ExecuteJSR(Memory() As AssemblerInstruction, Registers() As Integer, Address As Integer)
        Dim StackPointer As Integer = Registers(TOS) - 1 ' grabs next available location for top of memory (stack)
        Memory(StackPointer).OperandValue = Registers(PC) ' sets on stack (top of memory available) the operand value for the return address
        Registers(PC) = Address                           ' preforms the jump 
        Registers(TOS) = StackPointer                     ' sets up TOS for next jump
        DisplayStack(Memory, Registers)                   ' shows stacks (return values)
    End Sub

    ''' <summary>
    ''' Returns to lines specified in the top of Memory() from the Register(TOS) upward each time called.
    ''' </summary>
    ''' <param name="Memory"></param>
    ''' <param name="Registers"></param>
    Sub ExecuteRTN(Memory() As AssemblerInstruction, Registers() As Integer)
        Dim StackPointer As Integer = Registers(TOS)    ' grabs most recent stack
        Registers(TOS) += 1                             ' updates top of stack to go up one (to next stack)
        Registers(PC) = Memory(StackPointer).OperandValue ' performs jump to the line on the bottom of the stack/ most recent stack.
    End Sub


    ''' <summary>
    ''' executes the code.
    ''' </summary>
    ''' <param name="SourceCode"></param>
    ''' <param name="Memory"></param>
    Sub Execute(SourceCode() As String, Memory() As AssemblerInstruction)
        Dim Registers() As Integer = {0, 0, 0, 0, 0} ' REGISTTERS ARRAY. There are a few registers. View consts.
        SetFlags(Registers(ACC), Registers)  ' sets status red Z to 1 (100 -> ZNV)
        Registers(PC) = 0 ' initialises pc register to be  0. First instruction on line 0.
        Registers(TOS) = HI_MEM ' 4th register (Top Of Stack) is high men (max line count)

        '------------------------ displays useful info ------------------------
        Dim FrameNumber As Integer = 0
        DisplayFrameDelimiter(FrameNumber)
        DisplayCurrentState(SourceCode, Memory, Registers)
        '------------------------ displays useful info ------------------------

        Dim OpCode As String = Memory(Registers(PC)).OpCode  ' gets opcode of current instruction, dictated by pc. This is either JMP to START or JMP to line 1
        While OpCode <> "HLT"     ' HLT dictates the end of a program

            '------------------------ displays more useful info ------------------------
            FrameNumber += 1
            Console.WriteLine()
            DisplayFrameDelimiter(FrameNumber)
            '------------------------ displays more useful info ------------------------

            Dim Operand As Integer = Memory(Registers(PC)).OperandValue ' grabs operand value for the same opcode
            Console.WriteLine($"*  Current Instruction Register:  {OpCode} {Operand}") ' displays current insruction. i.e: ADD NUM1 (num1 operandvalue points to concrete value)
            Registers(PC) += 1 ' Increments pc by 1 
            Select Case OpCode
                Case "LDA"
                    ExecuteLDA(Memory, Registers, Operand)
                Case "STA"
                    ExecuteSTA(Memory, Registers, Operand)
                Case "LDA#"
                    ExecuteLDAimm(Registers, Operand)
                Case "ADD"
                    ExecuteADD(Memory, Registers, Operand)
                Case "JMP"
                    ExecuteJMP(Registers, Operand)
                Case "JSR"
                    ExecuteJSR(Memory, Registers, Operand)
                Case "CMP#"
                    ExecuteCMPimm(Registers, Operand)
                Case "BEQ"
                    ExecuteBEQ(Registers, Operand)
                Case "SUB"
                    ExecuteSUB(Memory, Registers, Operand)
                Case "SKP"
                    ExecuteSKP()
                Case "RTN"
                    ExecuteRTN(Memory, Registers)
            End Select
            If Registers(ERR) = 0 Then
                OpCode = Memory(Registers(PC)).OpCode                 ' grabs opcode for next instruction
                DisplayCurrentState(SourceCode, Memory, Registers)
            Else
                OpCode = "HLT"
            End If
        End While
        Console.WriteLine("Execution terminated")
    End Sub

    Sub AssemblerSimulator()
        Dim SourceCode(HI_MEM - 1) As String               ' string array 20 spaces big
        Dim Memory(HI_MEM - 1) As AssemblerInstruction     ' memory array. Each item contains 3 fields. opcode and operand string and value
        ResetSourceCode(SourceCode)
        ResetMemory(Memory)                                ' initialises Memory and SourceCode
        Dim Finished As Boolean = False
        Dim MenuOption As Char  ' input vector
        While Not Finished
            DisplayMenu()
            MenuOption = GetMenuOption()
            Select Case MenuOption
                Case "L"
                    LoadFile(SourceCode)
                    ResetMemory(Memory)
                Case "D"
                    If SourceCode(0) = EMPTY_STRING Then
                        Console.WriteLine("Error Code 7")
                    Else
                        DisplaySourceCode(SourceCode)
                    End If
                Case "E"
                    If SourceCode(0) = EMPTY_STRING Then
                        Console.WriteLine("Error Code 8")
                    Else
                        EditSourceCode(SourceCode)
                        ResetMemory(Memory)
                    End If
                Case "A"
                    If SourceCode(0) = EMPTY_STRING Then
                        Console.WriteLine("Error Code 9")
                    Else
                        Assemble(SourceCode, Memory)
                    End If
                Case "R"
                    If Memory(0).OperandValue = 0 Then
                        Console.WriteLine("Error Code 10")
                    ElseIf Memory(0).OpCode = "ERR" Then
                        Console.WriteLine("Error Code 11")
                    Else
                        Execute(SourceCode, Memory)
                    End If
                Case "X"
                    Finished = True
                Case Else
                    Console.WriteLine("You did not choose a valid menu option. Try again")
            End Select
        End While
        Console.WriteLine("You have chosen to exit the program")
        Console.ReadLine()
    End Sub

    Sub Main()
        AssemblerSimulator()
    End Sub
End Module
