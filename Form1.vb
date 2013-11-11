Public Class frmGame

    '###################################################################
    '
    '                       CLASS-SCOPE ARRAYS
    '
    '###################################################################

    'Declares a class-scope array of the type button
    Private _btnKeyboardArray(game.keyboardKeys.ARRAY_SIZE) As Button


    'Declares a class-scope array of the type string
    Private _sWordlist(game.wordlist.ARRAY_SIZE) As String


    'Declares a class-scope array of the type string
    Private _sHintlist(game.wordlist.ARRAY_SIZE) As String


    'Declares a class-scope array of the type Boolean
    Private _bHasGuessedLetter(game.wordlist.MAX_WORD_LENGTH) As Boolean


    'Declares a class-scope array of the type string
    Private _iPreviousUsedWords(game.MIN_GAMES_BEFORE_WORD_REUSE - 1) As Integer


    'Declares a class-scope array of the type Object
    Private _shpHangmanParts(game.MAX_WRONG_GUESSES - 1) As Object




    '###################################################################
    '
    '                       CLASS-SCOPE VARIABLES
    '
    '###################################################################

    Private _sCurrentWord As String                 'Declare a new private string
    Private _sCurrentHint As String                 'Declare a new private string
    Private _sGuessedLettersText As String          'Declare a new private string
    Private _sPlayerName As String                  'Declare a new private string

    Private _iTotalGuesses As Integer = 0           'Declare a new private integer
    Private _iWrongGuesses As Integer = 0           'Declare a new private integer
    Private _iGamesWon As Integer = 0               'Declare a new private integer
    Private _iGamesLost As Integer = 0              'Declare a new private integer
    Private _iHangmanPartsDisplayed As Integer = 0  'Declare a new private integer

    Private _bDisplayFullWord As Boolean = False    'Declare a new private boolean




#Region "Event Handlers"

    '###################################################################
    '
    '                       APPLICATION EVENT HANDLERS
    '
    '###################################################################

    Private Sub frmGame_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Declares a new boolean which will equal True if all files exist or false if files are missing
        Dim bAllFilesExist As Boolean = locateAllFiles()

        'If [bAllFilesExist] equals True, then all files exist - now fully load the program
        If bAllFilesExist = True Then

            Me.Text = getGameString(game.STRINGS_APP_TITLE)     'Set the title text for this form 
            Me.Size = New Size(game.APP_WIDTH, game.APP_HEIGHT) 'Set the size of this form

            Call randomizeNumberGeneration()    'Randomize number generation to prevent the same number occuring on startup
            Call prepareAllButtons()            'Creates all keyboard buttons and also sets the style of all buttons
            Call addWordsAndHints()             'Loads words and hints from an external text file into a global array
            Call setAllToolTips()               'Sets the tooltips for all objects(except the keyboard buttons) on the form
            Call populateHangmanPartsArray()    'Sets the value of [_shpHangmanParts] to contain the shape objects in the order they will appear
            Call loadFullGame()                 'Call loadFullGame to prepare application for a game

        Else    'If [bAllFilesExist] is false, then files are missing and the program will exit to prevent errors

            Call quickQuitApplication()  'Calls quitApplication function to quit the application

        End If  'End of if statement

    End Sub

    Private Sub frmGame_Shown(sender As System.Object, e As System.EventArgs) Handles MyBase.Shown
        Call showNewGameButton()    'Call showNeGameButton to display the new game button and focus it
    End Sub

    Private Sub btnNewGame_Click(sender As System.Object, e As System.EventArgs) Handles btnNewGame.Click
        Call newGame()      'Call newGame function to start a new game
    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Call loadFullGame() 'Call loadFullGame function to re-defined all game variables to defaults
    End Sub

    Private Sub btnQuit_Click(sender As System.Object, e As System.EventArgs) Handles btnQuit.Click
        Call quitApplication()  'Quit the application
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call loadFullGame() 'Call loadFullGame function to re-defined all game variables to defaults
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Dim sHelpText As String = getGameString(game.STRINGS_HELP)      'Declare a string for the help text
        Dim sTitle As String = getGameString(game.STRINGS_APP_TITLE)    'Declare a string for the title text
        MessageBox.Show(sHelpText, sTitle)                              'Display the messagebox with the help and title
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles QuitToolStripMenuItem.Click
        Call quitApplication()  'Quit the application
    End Sub

    Private Sub keyboardButton_Click(sender As System.Object, e As System.EventArgs)
        'Get the char guessed by removing the any prefix text from the button.text value
        Dim cLetterToGuess As Char = (Replace(sender.Text, game.KEYBOARD_BUTTON_TEXT_PREFIX, Nothing))

        Call disableKeyboardButton(sender)  'Disable the sender (the button which was clicked)
        Call guessLetter(cLetterToGuess)    'Call guessLetter function with the parameter of [cLetterToGuess]
    End Sub

#End Region




#Region "Game functions"

    '###################################################################
    '
    '                       GAME FUNCTIONS
    '
    '###################################################################


    'Called to process user guesses
    Private Sub guessLetter(cGuess As Char)

        Dim bAllLettersGuessed As Boolean = True
        'Declare a new local boolean and set the value to true


        'If the current word contains the guess letter then the guess was correct
        If _sCurrentWord.Contains(cGuess.ToString()) = True Then

            'The users guess was correct
            'now locate which position the correct character was at in the string

            'Loop through all of the chars in [_sCurrentWord]
            For iCounter As Integer = 0 To (_sCurrentWord.Length - 1)

                'If the current char is equal to the user's guess
                If _sCurrentWord(iCounter) = cGuess.ToString Then

                    'Set the value of [_bHasGuessedLetter] at the same index to True
                    _bHasGuessedLetter(iCounter) = True


                    'Else if the char is not equal to the user's guess but is equal to the SPACE char then
                ElseIf _sCurrentWord(iCounter) = game.SPACE_CHAR Then

                    'Treat spaces as correct guesses as there is nothing to guess
                    'Set the value of [_bHasGuessedLetter] at the same index to True
                    _bHasGuessedLetter(iCounter) = True

                End If   'End of IF statement


                'If the index in [_bHasGuessLetter] is still false then
                'the user has not yet correctly guessed this letter and also did not guess the letter
                'this round - therefore the user has not guessed all of the letters correctly
                If _bHasGuessedLetter(iCounter) = False Then
                    bAllLettersGuessed = False      'Set [bAllLettersGuessed] to false
                End If

            Next    'Continue onto the next loop

        Else 'If the current word does not contain the user's guess - then the user has made an incorrect guess

            _iWrongGuesses += 1             'Increment the value of [_iWrongGuesses] by 1
            bAllLettersGuessed = False      'Set [bAllLettersGuessed] to false
            Call displayNextHangmanPart()   'Display the next hangman part 

        End If 'End of IF statement


        'If [_sGuessedLettersText] is nothing then this is the users first guess this round
        If _sGuessedLettersText = Nothing Then

            _sGuessedLettersText = cGuess.ToString()
            'Set the value of [_sGuessedLettersText] to the letter guessed

        Else

            _sGuessedLettersText = _sGuessedLettersText & game.LIST_SEPARATOR_CHAR & cGuess.ToString()
            'Append the list separator char (a comma), followed by the users guess

        End If      'End of IF statement


        _iTotalGuesses += 1         'Increment the number of total guesses by 1
        Call updateGameLabels()     'Call function to update the labels on the form


        'If bAllLettersGuessed is still true at this point
        'Then the user has correctly guessed every character in the current word
        'This means they have won the game
        If bAllLettersGuessed = True Then
            Call gameWon()      'Call gameWon()


            'Elseif the number of wrong guesses is greater than or equal to
            'the max amount of wrong guesses - the user has lost the game 
        ElseIf _iWrongGuesses >= game.MAX_WRONG_GUESSES Then
            Call gameLost()         'Call gameLost()

        End If                      'End of IF statement

    End Sub

    Private Sub newGame()
        Dim iGeneratedNumber As Integer = getNextWordPosition()
        'Declare a new integer and set the value to the value returned by getNextWordPosition()


        'Reset the word display to prevent the new answer being displayed due to 
        '[_bDisplayFullWord] being equal to true from the last game finish
        Call resetWordDisplay()

        Call resetCurrentValues()   'Call resetCurrentValues()
        _sCurrentWord = _sWordlist(iGeneratedNumber)    'Set [_sCurrentWord] to the new word
        _sCurrentHint = _sHintlist(iGeneratedNumber)    'Set [_sCurrentHint] to the new hint
        Call updateGameLabels()     'Update form label values

        Call enableAllKeyboardButtons()         'Enable all keyboard buttons
        Call hideHangman()                      'Hide the hangman
        Call hideNewGameButton()                'Hide the newgame button

        lblEndGameSubtitle.Visible = False      'Ensure end game subtitle is not visible
        lblEndGameTitle.Visible = False         'Ensure end game title is not visible
    End Sub

    Private Sub gameWon()
        _iGamesWon += 1     'Increment games won by 1

        'Set the ForeColor of the end game title and subtitle to the winner colour
        lblEndGameTitle.ForeColor = game.COLOURS_WINNER_TEXT
        lblEndGameSubtitle.ForeColor = game.COLOURS_WINNER_TEXT


        'Set the text of the end game title and subtitle to the winner title and subtitle text
        'The text for these values are stored as a constant in the 'game' class
        'The subtitle has a variable placeholder which is replaced by the value of [_sPlayerName]
        lblEndGameTitle.Text = getGameString(game.STRINGS_WINNER_TITLE)
        lblEndGameSubtitle.Text = getGameString(game.STRINGS_WINNER_SUBTITLE, _sPlayerName)

        Call gameEnd()      'Call game end function
    End Sub

    Private Sub gameLost()
        _iGamesLost += 1    'Increment games won by 1

        'Set the ForeColor of the end game title and subtitle to the loser colour
        lblEndGameTitle.ForeColor = game.COLOURS_LOSER_TEXT
        lblEndGameSubtitle.ForeColor = game.COLOURS_LOSER_TEXT


        'Set the text of the end game title and subtitle to the loser title and subtitle text
        'The text for these values are stored as a constant in the 'game' class
        'The subtitle has a variable placeholder which is replaced by the value of [_sPlayerName]
        lblEndGameTitle.Text = getGameString(game.STRINGS_LOSER_TITLE)
        lblEndGameSubtitle.Text = getGameString(game.STRINGS_LOSER_SUBTITLE, _sPlayerName)

        Call gameEnd()      'Call end game function
    End Sub

    'This function is called every time a game ends - independant of if the game was lost or won
    Private Sub gameEnd()
        Call disableAllKeyboardButtons()    'Disable all keyboard buttons
        Call highlightAndFillWordDisplay()  'Call function to highligh and fill the word display with the correct answer
        Call showNewGameButton()            'Call function to show the new game button

        lblEndGameSubtitle.Visible = True   'Set the end game subtitle to be visible
        lblEndGameTitle.Visible = True      'Set the end game title to be visible
    End Sub

    Private Sub highlightAndFillWordDisplay()
        lblWordDisplay.ForeColor = game.COLOURS_WORD_HIGHLIGHTED    'Set lblWordDisplay's ForeColor to the the highlighted value
        _bDisplayFullWord = True    'Set _bDisplayFullWord to true to cause the full word to display on label update
        Call updateGameLabels()     'Update the labels on the form to have the correct word displayed fully
    End Sub

    Private Sub resetWordDisplay()
        lblWordDisplay.ForeColor = game.COLOURS_WORD_NORMAL 'Set lblWordDisplay's ForeColor to the the default value
        _bDisplayFullWord = False   'Set _bDisplayFullWord to false to stop the word displaying fully
        Call updateGameLabels()     'Update the form labels
    End Sub

    Private Sub updateGameLabels()
        Dim sWordText As String = ""    'Declare a new empty string

        If _sCurrentWord IsNot Nothing Then 'If the current word is not Nothing then

            'Start a loop through the current word to create the word display for the user
            'Underscores will replace letters which have not been guessed
            For iCounter As Integer = 0 To (_sCurrentWord.Length - 1)

                'If the current character is a space - append the word separator string
                If _sCurrentWord(iCounter) = game.SPACE_CHAR Then
                    sWordText = sWordText & game.WORD_SEPARATOR


                    'If the user has already correctly guessed this letter or 
                    '_bDisplayFullWord equals true then -
                    'append the letter separator string and then the current character
                ElseIf (_bHasGuessedLetter(iCounter) = True) Or (_bDisplayFullWord = True) Then
                    sWordText = sWordText & game.LETTER_SEPARATOR & _sCurrentWord(iCounter)


                    'If the user has not guessed the letter and the letter is not a space
                    'Then append the letter separator string and an underscore to act
                    'as a placeholder for the unknown letter
                Else
                    sWordText = sWordText & game.LETTER_SEPARATOR & game.UNKNOWN_LETTER_CHAR


                End If          'End of IF statement within the loop
            Next                'End of Loop

        End If                  'End of IF Statement


        'Set the text of lblHint to the constant STRINGS_HINT, 
        'replacing the 1st variable place holder with the value of [_sCurrentHint]
        'replacing the 2nd variable place holder with the value of [_sCurrentWord.Length]
        lblHint.Text = getGameString(game.STRINGS_HINT, _sCurrentHint, _sCurrentWord.Length)


        'Set the text of lblLetterGuessed to the constant STRINGS_LETTERS_GUESSED
        'replacing the variable place holder with the value of [_sGuessedLettersText]
        lblLettersGuessed.Text = getGameString(game.STRINGS_LETTERS_GUESSED, _sGuessedLettersText)


        'Sets the text of lblWordDisplay to the value of sWordText
        lblWordDisplay.Text = sWordText


        'Set the text of lblScore to the constant STRINGS_SCORE
        'replacing the 1st variable place holder with the value returned by getUserWinPercentage()
        'replacing the 2st variable place holder with the value of [_iGamesWon]
        'replacing the 1st variable place holder with the value of [_iGamesLost]
        lblScore.Text = getGameString(game.STRINGS_SCORE, getUserWinPercentage(), _iGamesWon, _iGamesLost)


        'Set the text of lblAmountOfGuesses to the value of the constant STRINGS_WRONG_GUESSES_AMOUNT
        'replacing the 1st varible place holder with the value of [_iWrongGuesses] and
        'replacing the 2nd variable place holder with the value of the constant MAX_WRONG_GUESSES
        lblAmountOfGuesses.Text = getGameString(game.STRINGS_WRONG_GUESSES_AMOUNT, _iWrongGuesses, game.MAX_WRONG_GUESSES)
    End Sub

    Private Function getUserWinPercentage() As Double
        'Declares a new integer equal to the value of games won plus games lost 
        Dim iTotalGamesPlayed As Integer = _iGamesWon + _iGamesLost

        'Declares a new double as games won divided by the games played - multipled by 100
        Dim iUserWinPercentage As Double = ((_iGamesWon / iTotalGamesPlayed) * 100)

        'If [iUserWinPercentage] is not a number (NaN) then either games won or total games played
        'are equal to 0 - either way, the user has a 0 win percentage
        If (Double.IsNaN(iUserWinPercentage) = True) Then
            iUserWinPercentage = 0.0 'Manual set the win percentage to 0
        End If


        'Round to win percentage to the the number of decimal places stated in 
        'the constant SCORE_PRECISION_AFTER_DECIMAL_POINT in the 'game' class
        iUserWinPercentage = Math.Round(iUserWinPercentage, game.SCORE_PRECISION_AFTER_DECIMAL_POINT)

        Return iUserWinPercentage   'Return the win percentage
    End Function

    Private Sub displayNextHangmanPart()
        'If the number of hangman parts displayed does not exceed the range of the hangman parts array then
        If (_iHangmanPartsDisplayed) <= (_shpHangmanParts.Length - 1) Then

            'Set the next hangman part to become visible
            _shpHangmanParts(_iHangmanPartsDisplayed).Visible = True

        End If          'End of IF statement

        _iHangmanPartsDisplayed += 1    'Increment the value of [_iHangmanPartsDisplayed] by 1
    End Sub

    Private Sub hideHangman()
        'Loop through all of the hangman parts array and make all of the objects invisible
        For iCounter As Integer = 0 To (_shpHangmanParts.Length - 1)
            _shpHangmanParts(iCounter).Visible = False
        Next

        _iHangmanPartsDisplayed = 0     'Set the value of [_iHangmanPartsDisplayed] to 0
    End Sub

#End Region




#Region "Wordlist functions"

    '###################################################################
    '
    '       EXTERNAL WORDLIST FILE READING AND STORING IN MEMORY
    '
    '###################################################################

    Private Sub addWordsAndHints()
        Dim sCurrentLine As String                      'Declare a new local string
        Dim sWord As String                             'Declare a new local string
        Dim sHint As String                             'Declare a new local string
        Dim iWordsLoaded As Integer = 0                 'Declare a new local integer
        Dim iWordsSkippedDueToLength As Integer = 0     'Delcare a new local integer
        Dim sWordsAndHintsArray() As String             'Declare a new string array

        'Declare a new string a set it to the value of the base path of the applicaitons exe file
        'and the value of the constant WORDLIST_PATH
        Dim sFullFilePath As String = game.FILES_BASE_PATH & game.FILES_WORDLIST_PATH


        'Declare a new string array and set it to the file contents of the file location at [sFullFilePath]
        'split upon each new line - with one line per array position
        Dim sFileLinesArray() As String = getArrayFromTextFile(sFullFilePath, game.LINE_BREAK)


        'Loop through each line in [sFileLinesArray]
        For Each sCurrentLine In sFileLinesArray

            'Set sWordsAndHintsArray to sCurrentLine split by the value of WORDLIST_SPLIT_CHAR
            sWordsAndHintsArray = Split(sCurrentLine, game.WORDLIST_SPLIT_CHAR)

            'Set sWord to the word's position in the line & set sHint to the hint's position in the line
            'Both sWord and sHint are set ToUpper as the character input is also ToUpper
            sWord = sWordsAndHintsArray(game.wordlist.WORD_POSITION_IN_LINE_ARRAY).ToUpper()
            sHint = sWordsAndHintsArray(game.wordlist.HINT_POSITION_IN_LINE_ARRAY).ToUpper()


            'If the word length is less than or equal to the MAX_WORD_LENGTH and
            'the hint's length is less than or equal to the MAX_HINT_LENGTH
            'then add the word and it's hint to the arrays
            If (sWord.Length <= game.wordlist.MAX_WORD_LENGTH) AndAlso (sHint.Length <= game.wordlist.MAX_HINT_LENGTH) Then
                _sWordlist(iWordsLoaded) = sWord    'Add the word in the next free position in the word list
                _sHintlist(iWordsLoaded) = sHint    'Add the hint in the next free position in the hint list
                iWordsLoaded += 1                   'Increment words loaded by 1
            Else
                iWordsSkippedDueToLength += 1       'Increment the number of skipped words by 1
            End If
        Next

        ReDim Preserve _sWordlist(0 To (iWordsLoaded - 1))  'Re-declare the size of sWordList to the number of words added
        ReDim Preserve _sHintlist(0 To (iWordsLoaded - 1))  'Re-declare the size of sHintLint to the number of hints added

        'If any words have been skipped then display a message indicating the amount of words skipped to the user
        If (iWordsSkippedDueToLength > 0) Then

            'Declare a new string and set the value to the STRINGS_WORDS_REJECTED constant
            'Replace variable place holder with the amount of words skipped
            Dim skippedWordsText As String = getGameString(game.STRINGS_WORDS_REJECTED, iWordsSkippedDueToLength.ToString)

            'Display the message to the user
            MsgBox(skippedWordsText)
        End If

    End Sub

    Private Function getNextWordPosition()
        'Declare a new integer and set the value to a random number between 0 and sWordList's maximum position
        Dim iGeneratedNumber As Integer = getRandomNumber(0, (_sWordlist.Length - 1))

        'Declare a new boolean and set it to the value returned by canUseGeneratedNumber
        Dim bCanUseNumber As Boolean = canUseGeneratedWord(iGeneratedNumber)

        'Declare an integer and set the value to 0
        Dim iGenerationAttempts As Integer = 0


        'If [bCanUseNumber] is false then the number cannot be used as it has recently been used 
        'and another must be generated.
        'the number of generation attempts must also be less than the value of the MAX_GENERATION_ATTEMPTS constant.
        '[iGeneratonAttempts] is used to prevent an infinite loop occuring when an acceptable
        'word cannot be found within the wordlist.
        While ((bCanUseNumber = False) And (iGenerationAttempts < game.MAX_GENERATION_ATTEMPTS))

            iGeneratedNumber = getRandomNumber(0, (_sWordlist.Length - 1))  'Generate a new number
            bCanUseNumber = canUseGeneratedWord(iGeneratedNumber)           'Check if the new number can be used
            iGenerationAttempts += 1                                        'Increment the amount of attempts by 1

        End While                                                           'End of loop

        'The value of [iGeneratedNumber] is now the accepted number
        'Update the previous used words list with the new word
        Call updatePreviousUsedWords(iGeneratedNumber)

        Return iGeneratedNumber                                             'Return the decided position in the word list
    End Function

    Private Function canUseGeneratedWord(iGeneratedNumber As Integer)
        Dim bCanUse As Boolean = True   'Declares a boolean and sets it to True


        'Loops through [_iPreviousUsedWords] and checks if [iGeneratedNumber] is already
        'stored inside the array - in which case the word has been used recently and
        'a new word should be used instead
        For iCounter As Integer = 0 To (_iPreviousUsedWords.Length - 1)

            If iGeneratedNumber = _iPreviousUsedWords(iCounter) Then 'If [iGeneratedNumber] is within the [_iPreviousUsedWords] array
                bCanUse = False                                      'Set [bCanUse] to false
                Exit For                                             'Exit the loop once the number has been found in the array
            End If                                                   'End of IF statement

        Next


        Return bCanUse                                               'Return the value of [bCanUse]
    End Function

    Private Sub updatePreviousUsedWords(iNewIndex As Integer)
        'Loops through [_iPreviousUsedWords] from the last value in the array
        'to the 2nd first value in the array - setting each array position to the value
        'of the position before it - effectively moving every value in the array back one place
        'and removing the oldest value from the array
        For iCounter As Integer = (_iPreviousUsedWords.Length - 1) To 1 Step -1
            _iPreviousUsedWords(iCounter) = _iPreviousUsedWords(iCounter - 1)
        Next


        'Add the new word index to the front of the array
        _iPreviousUsedWords(0) = iNewIndex
    End Sub
#End Region




#Region "Button Functions"


    '###################################################################
    '
    '                   BUTTION RELATED FUNCTIONS/SUBS
    '
    '###################################################################

    Private Sub prepareAllButtons()
        Call addKeyboard()                  'Adds all of the keyboard buttons

        Call setButtonStyle(btnNewGame)     'Changes various properties of btnNewGame
        Call setButtonStyle(btnQuit)        'Changes various properties of btnQuit
        Call setButtonStyle(btnReset)       'Changes various properties of btnReset


        'Set a custom BackColor for btnNewGame to make it stand out
        btnNewGame.BackColor = game.COLOURS_BUTTON_NEW_GAME_BACK
    End Sub

    Private Sub showNewGameButton()
        btnNewGame.Visible = True   'Make btnNewGame Visible
        btnNewGame.Enabled = True   'Enable btnNewGame
        btnNewGame.Focus()          'Focus btnNewGame
    End Sub

    Private Sub hideNewGameButton()
        btnNewGame.Visible = False  'Make btnNewGame Invisible
        btnNewGame.Enabled = False  'Disable btnNewGame

    End Sub

    Private Sub addKeyboard()
        'Declare a new string array and set it to thee value of 
        'KEYBOARD_TEXT split by the value of KEYBOARD_SPLIT_CHAR
        Dim sLetterArray() As String = Split(game.KEYBOARD_TEXT, game.KEYBOARD_SPLIT_CHAR)

        Dim iCurrentY As Integer = game.keyboardKeys.Y  'Declare a new integer
        Dim iCurrentX As Integer = game.keyboardKeys.X  'Declare a new integer
        Dim iRowsAdded As Integer = 0                   'Declare a new integer
        Dim iButtonIndex As Integer = 0                 'Declare a new integer


        'Loop through each value within [sLetterArray]
        For iCounter As Integer = 0 To (sLetterArray.Length - 1)

            'If the current character is equal to KEYBOARD_NEW_ROW_CHAR then
            'take a new row rather than adding a button with that character
            If sLetterArray(iCounter) = game.KEYBOARD_NEW_ROW_CHAR Then

                iRowsAdded += 1     'Increment the number of rows

                'Set X and Y integers to the default value + any increments times the amount of rows added
                iCurrentY = game.keyboardKeys.Y + (iRowsAdded * game.keyboardKeys.Y_INCREMENT)
                iCurrentX = game.keyboardKeys.X + (iRowsAdded * game.keyboardKeys.ROW_INCREMENT)


                'If the character was not the new row char then create a button
                'with the character as the buttons text.
            Else

                iButtonIndex = (iCounter - iRowsAdded)  'Prevent blank spaces occuring in the button array


                'Create a new button and set all necessary properties for the button
                _btnKeyboardArray(iButtonIndex) = New Button()
                _btnKeyboardArray(iButtonIndex).Name = game.KEYBOARD_BUTTON_NAME_PREFIX & sLetterArray(iCounter)
                _btnKeyboardArray(iButtonIndex).Text = game.KEYBOARD_BUTTON_TEXT_PREFIX & sLetterArray(iCounter)
                _btnKeyboardArray(iButtonIndex).Size = New Size(game.keyboardKeys.BUTTON_WIDTH, game.keyboardKeys.BUTTON_HEIGHT)
                _btnKeyboardArray(iButtonIndex).Location = New Point(iCurrentX, iCurrentY)
                _btnKeyboardArray(iButtonIndex).Enabled = True
                _btnKeyboardArray(iButtonIndex).Visible = True


                'Set the visual properties of the newly created button with the setButtonStyle function
                Call setButtonStyle(_btnKeyboardArray(iButtonIndex))

                'Set the tooltip for the button and add it to the form's controls
                tltTip.SetToolTip(_btnKeyboardArray(iButtonIndex), getGameString(game.STRINGS_KEYBOARD_TOOLTIP, sLetterArray(iCounter)))
                Me.Controls.Add(_btnKeyboardArray(iButtonIndex))

                'Add an event handler for the click event of the button
                AddHandler _btnKeyboardArray(iButtonIndex).Click, AddressOf keyboardButton_Click

                'Increment currentX by the value of X_INCREMENT
                iCurrentX += game.keyboardKeys.X_INCREMENT
            End If
        Next
    End Sub

    Private Sub setButtonStyle(ByRef btnName As Button)
        'Set tabStop property of button to false
        'This prevents buttons from becoming automatically focused
        btnName.TabStop = False

        'Set the button's style to Flat
        btnName.FlatStyle = FlatStyle.Flat

        'Set the button's border colour and size
        btnName.FlatAppearance.BorderColor = game.COLOURS_BUTTON_BORDER
        btnName.FlatAppearance.BorderSize = game.keyboardKeys.BORDER_SIZE

        'Set the mouse over backcolour and mouse down backcolour for the button
        btnName.FlatAppearance.MouseDownBackColor = game.COLOURS_BUTTON_MOUSEDOWN
        btnName.FlatAppearance.MouseOverBackColor = game.COLOURS_BUTTON_MOUSEOVER

        'Set the forecolor of the button
        btnName.ForeColor = game.COLOURS_BUTTON_FORE
    End Sub

    Private Sub enableAllKeyboardButtons()
        'Call enableKeyboardButton on each object within [_btnKeyboardArray]
        For iCounter As Integer = 0 To (_btnKeyboardArray.Length - 1)
            Call enableKeyboardButton(_btnKeyboardArray(iCounter))
        Next
    End Sub

    Private Sub disableAllKeyboardButtons()
        'Call disableKeyboardButton on each object within [_btnKeyboardArray]
        For iCounter As Integer = 0 To (_btnKeyboardArray.Length - 1)
            Call disableKeyboardButton(_btnKeyboardArray(iCounter))
        Next
    End Sub

    Private Sub disableKeyboardButton(ByRef btnToDisable As Button)
        'Disable the button and set its borderColor to the disable border colour
        btnToDisable.Enabled = False
        btnToDisable.FlatAppearance.BorderColor = game.COLOURS_BUTTON_DISABLED_BORDER
    End Sub

    Private Sub enableKeyboardButton(ByRef btnToEnable As Button)
        'Enable the button and set its borderColor to the normal border colour
        btnToEnable.Enabled = True
        btnToEnable.FlatAppearance.BorderColor = game.COLOURS_BUTTON_BORDER
    End Sub
#End Region




#Region "Core functions"

    Private Function getPlayerName()
        'Set sName to the value input by the user into an inputbox
        'The inputBox displays the value of STRINGS_ENTER_NAME as instructions to the user
        'The value of STRINGS_APP_TITLE is used as the inputBox title
        Dim sName As String = InputBox(getGameString(game.STRINGS_ENTER_NAME), getGameString(game.STRINGS_APP_TITLE))

        'If the user did not enter a value into the inputbox, then give then a default name
        'which is stored in the constant STRINGS_DEFAULT_NAME
        If sName = Nothing Then
            sName = getGameString(game.STRINGS_DEFAULT_NAME)
        End If

        'Return the value of sName
        Return sName
    End Function

    Private Function locateAllFiles()
        Dim bAllFilesExist As Boolean = True    'Declare a new boolean as true
        Dim iFilesMissing As Integer = 0        'Declare a new integer and set it to 0


        'Loop through FILES_ALL_PATHS - which is located in the 'game' class - and check if 
        'the each value leads to an exisiting file. If a file is missing then set bAllFilesExist
        'to false and increment the value of iFilesMissing by 1
        For Each sFilePath As String In game.FILES_ALL_PATHS
            If My.Computer.FileSystem.FileExists(sFilePath) = False Then
                bAllFilesExist = False
                iFilesMissing += 1
            End If
        Next

        'If bAllFilesEixst is false then some files which should exist, currently do not.
        'Output an error message stored in STRINGS_UNABLE_TO_LOCATE_FILES with the variable place holder
        'replaced with the amount of files missing. The message box is given the title of STRINGS_APP_TITLE
        If bAllFilesExist = False Then
            Dim sErrorMessage As String = getGameString(game.STRINGS_UNABLE_TO_LOCATE_FILES, iFilesMissing.ToString)
            MessageBox.Show(sErrorMessage, getGameString(game.STRINGS_APP_TITLE))
        End If


        'Return the value of bAllFilesExist
        Return bAllFilesExist
    End Function

    Private Sub quickQuitApplication()
        Me.Close()                      'Close the application
    End Sub

    Private Sub quitApplication()
        'Declare two strings with the display text and title for the message box asking if the user is sure
        'they wish to quit the application
        Dim sAreYouSureText As String = getGameString(game.STRINGS_QUIT_ARE_YOU_SURE)
        Dim sTitle As String = getGameString(game.STRINGS_APP_TITLE)

        'Display the messagebox and show the result within [drIsSure]
        Dim drIsSure As DialogResult = MessageBox.Show(sAreYouSureText, sTitle, MessageBoxButtons.YesNo)

        'If the value of [drIsSure] is equal to DialogResult.Yes then the user selected Yes
        'on the message box and wants to quit - so call quickQuitApplication to close the application
        If drIsSure = Windows.Forms.DialogResult.Yes Then
            Call quickQuitApplication()
        End If
    End Sub

    Private Function getArrayFromTextFile(sPath As String, sSplitOn As String)
        'Declare a new string and give it a default value from the constant STRINGS_DEFAULT_FILE_RETURN
        Dim sFileContents As String = game.STRINGS_DEFAULT_FILE_RETURN

        'Check if the file they wish to read exists before attempting to read it to prevent crashes.
        'If the file exists then read the contents and store them inside [sFileContents] and call
        'a function to filter unnecessary visual characters - such as TAB characters
        If My.Computer.FileSystem.FileExists(sPath) Then
            sFileContents = My.Computer.FileSystem.ReadAllText(sPath)
            sFileContents = filterTextFileInput(sFileContents)
        End If

        'Create a new array and set it to the value of [sFileContents] split by the parameter [sSplitOn]
        'and then return the array
        Dim sFileContentsArray() As String = Split(sFileContents, sSplitOn)
        Return sFileContentsArray
    End Function

    Private Function getRandomNumber(iMinimum As Integer, iMaximum As Integer)
        'Returns a random number between [iMinimum] and [iMaximum]
        Return Int((iMaximum - iMinimum + 1) * Rnd() + iMinimum)
    End Function

    Private Sub randomizeNumberGeneration()
        'Randomizes number generation using the system time from the machine.
        'This prevents the same numbers occuring when the application is closed and re-opened
        Call Randomize()
    End Sub

    Private Function getGameString(ByVal sText As String, Optional sValueOne As String = Nothing, Optional sValueTwo As String = Nothing, Optional sValueThree As String = Nothing)
        'Declare a new string and set it to the value of sText
        'Then replace any instances of the value of STRINGS_NEW_LINE_PLACEHOLDER with a new line
        Dim sFinalText As String = sText
        sFinalText = Replace(sFinalText, game.STRINGS_NEW_LINE_PLACEHOLDER, game.LINE_BREAK)


        'If sValueX has been defined by the user then replace the Xth place holder with the value
        'input by the user. This is done for three place holders currently.
        If sValueThree IsNot Nothing Then
            sFinalText = Replace(sFinalText, game.STRINGS_3RD_VAR_PLACEHOLDER, sValueThree)
        End If

        If sValueTwo IsNot Nothing Then
            sFinalText = Replace(sFinalText, game.STRINGS_2ND_VAR_PLACEHOLDER, sValueTwo)
        End If

        If sValueOne IsNot Nothing Then
            sFinalText = Replace(sFinalText, game.STRINGS_1ST_VAR_PLACEHOLDER, sValueOne)
        End If


        'Return the value of sFinalText
        Return sFinalText
    End Function

    Private Function filterTextFileInput(fileContents As String)
        'Declare a new string and set it to the value of fileContents
        Dim sCleanText As String = fileContents

        'Replace TAB characters with nothing to remove them
        sCleanText = Replace(sCleanText, game.TAB_KEY, Nothing)


        'Return the text with unnecessary characters removed
        Return sCleanText
    End Function

    Private Sub loadFullGame()
        Call resetCurrentValues()           'Call resetCurrentValues to reset various game-to-game variables
        _sPlayerName = getPlayerName()      'Set [_sPlayerName] to the value returned by getPlayerName()

        _iTotalGuesses = 0                  'Set the number of guesses to 0
        _iGamesLost = 0                     'Set the number of games lost to 0
        _iGamesWon = 0                      'Set the number of games won to 0

        lblEndGameTitle.Visible = False     'Hide end game title label incase it is visible
        lblEndGameSubtitle.Visible = False  'Hide end game subtitle label incase it is visible
        Call hideHangman()                  'Call hideHangman to hide the hangman


        'Loop through all values in the previously guessed array
        For iCounter As Integer = 0 To (_iPreviousUsedWords.Length - 1)

            'Set all values to the default value - which is defined in a constant
            _iPreviousUsedWords(iCounter) = game.DEFAULT_PREVIOUS_GENERATED_NUMBER

        Next 'End of for loop

        Call disableAllKeyboardButtons()    'Call function to disable all keyboard buttons
        Call updateGameLabels()             'Update all the label object's values
        Call showNewGameButton()            'Show the new game button and focus it
    End Sub

    Private Sub resetCurrentValues()
        _sGuessedLettersText = ""   'Set the variable value to nothing
        _sCurrentWord = ""          'Set the variable value to nothing
        _sCurrentHint = ""          'Set the variable value to nothing
        _iWrongGuesses = 0          'Set the amount of wrong guesses to 0

        For counter As Integer = 0 To (_bHasGuessedLetter.Length - 1)   'Loop through the values of [_bHasGuessLetter]
            _bHasGuessedLetter(counter) = False         'Set all values to False
        Next                                            'End of for loop
    End Sub

    Private Sub setAllToolTips()
        tltTip.SetToolTip(btnNewGame, getGameString(game.STRINGS_NEW_GAME_TOOLTIP)) 'Set the tooltip for btnNewGame
        tltTip.SetToolTip(btnReset, getGameString(game.STRINGS_RESET_TOOLTIP))      'Set the tooltip for btnReset
        tltTip.SetToolTip(btnQuit, getGameString(game.STRINGS_QUIT_TOOLTIP))        'Set the tooltip for btnQuit
    End Sub

    Private Sub populateHangmanPartsArray()
        'Set the value of [_shpHangmanParts} to a list of all hangman objects in the order they will appear.
        'This array cannot be populated until the form has created the hangman objects or a null reference
        'error will occur - a suitable place to for this to be called is on form load.
        _shpHangmanParts = { _
                    shpHead, _
                    shpBody, _
                    shpLeftArm, _
                    shpRightArm, _
                    shpLeftLeg, _
                    shpRightLeg _
                            }
    End Sub
#End Region



End Class




Public NotInheritable Class game

    '###################################################################
    '
    '   THIS CLASS STORES GLOBAL-SCOPE:
    '                                   CONSTANTS
    '                                   ENUMS
    '                                   SHARED READONLY VARIABLES
    '
    '###################################################################


    Public Const FILES_WORDLIST_PATH As String = "\wordlist.txt"



    Public Const APP_HEIGHT As Integer = 600
    Public Const APP_WIDTH As Integer = 680
    Public Const MAX_WRONG_GUESSES As Integer = 6
    Public Const MAX_GENERATION_ATTEMPTS As Integer = 5
    Public Const MIN_GAMES_BEFORE_WORD_REUSE As Integer = 3
    Public Const DEFAULT_PREVIOUS_GENERATED_NUMBER As Integer = -1
    Public Const SCORE_PRECISION_AFTER_DECIMAL_POINT As Integer = 2

    Public Const LIST_SEPARATOR_CHAR As Char = ","c
    Public Const SPACE_CHAR As Char = " "c
    Public Const UNKNOWN_LETTER_CHAR As Char = "_"c
    Public Const WORD_SEPARATOR As String = "    "
    Public Const LETTER_SEPARATOR As String = " "
    Public Const LINE_BREAK As String = vbCrLf
    Public Const TAB_KEY As String = vbTab



    Public Const STRINGS_1ST_VAR_PLACEHOLDER As String = "#%1#"
    Public Const STRINGS_2ND_VAR_PLACEHOLDER As String = "#%2#"
    Public Const STRINGS_3RD_VAR_PLACEHOLDER As String = "#%3#"
    Public Const STRINGS_NEW_LINE_PLACEHOLDER As String = "\n "

    Public Const STRINGS_APP_TITLE As String = "Hangman game by David Watson"
    Public Const STRINGS_HELP As String = "Select the new game button to start a game and look at the hint at the top left. \n \n Click a button or press the key on your keyboard to guess a letter."
    Public Const STRINGS_ENTER_NAME As String = "Please enter your name!"
    Public Const STRINGS_UNABLE_TO_LOCATE_FILES As String = "Unable to locate #%1# file(s). This application will now close."
    Public Const STRINGS_WORDS_REJECTED As String = "#%1# word(s) have been ignored as they exceed the length limit"
    Public Const STRINGS_QUIT_ARE_YOU_SURE As String = "Are you sure you wish to quit the application?"

    Public Const STRINGS_DEFAULT_FILE_RETURN As String = "NULL"
    Public Const STRINGS_DEFAULT_NAME As String = "Anonymous"

    Public Const STRINGS_QUIT_TOOLTIP As String = "Press to quit the application"
    Public Const STRINGS_KEYBOARD_TOOLTIP As String = "Click to guess the letter #%1#"
    Public Const STRINGS_NEW_GAME_TOOLTIP As String = "Press to continue to the next game"
    Public Const STRINGS_RESET_TOOLTIP As String = "Press to reset the application variable"

    Public Const STRINGS_HINT As String = "HINT: #%1# \n \n Word Length: #%2#"
    Public Const STRINGS_SCORE As String = "Win Percentage: #%1# % \n Games Won:       #%2# \n Games Lost:        #%3#"
    Public Const STRINGS_LETTERS_GUESSED As String = "Letters guessed so far: \n #%1#"
    Public Const STRINGS_WRONG_GUESSES_AMOUNT As String = "Number of wrong guesses: #%1# / #%2#"
    Public Const STRINGS_LOSER_TITLE As String = "Tough Luck!"
    Public Const STRINGS_LOSER_SUBTITLE As String = "Better luck next time, #%1#!"
    Public Const STRINGS_WINNER_TITLE As String = "Congratulations!"
    Public Const STRINGS_WINNER_SUBTITLE As String = "Good job, #%1#!"




    Public Const KEYBOARD_TEXT As String = "Q W E R T Y U I O P | A S D F G H J K L | Z X C V B N M"
    Public Const KEYBOARD_SPLIT_CHAR As Char = " "c
    Public Const KEYBOARD_NEW_ROW_CHAR As Char = "|"c
    Public Const KEYBOARD_BUTTON_NAME_PREFIX As String = "btnKeyboard_"
    Public Const KEYBOARD_BUTTON_TEXT_PREFIX As String = "&"
    Public Enum keyboardKeys As Integer                                     'Create a Enum for keyboard-generation settings
        ARRAY_SIZE = 25
        X = 270
        Y = 450
        X_INCREMENT = 35
        Y_INCREMENT = 40
        ROW_INCREMENT = 15
        BUTTON_HEIGHT = 30
        BUTTON_WIDTH = 30
        BORDER_SIZE = 1
    End Enum




    Public Const WORDLIST_SPLIT_CHAR As Char = "|"c
    Public Enum wordlist As Integer                                         'Create a enum for wordlist information
        ARRAY_SIZE = 99
        INFORMATION_PER_LINE = 2
        WORD_POSITION_IN_LINE_ARRAY = 0
        HINT_POSITION_IN_LINE_ARRAY = 1
        MAX_WORD_LENGTH = 15
        MAX_HINT_LENGTH = 60
    End Enum



    Public Shared ReadOnly FILES_BASE_PATH As String = Application.StartupPath()
    Public Shared ReadOnly FILES_ALL_PATHS() As String = { _
                                                            FILES_BASE_PATH & FILES_WORDLIST_PATH _
                                                         }



    Public Shared ReadOnly COLOURS_BUTTON_FORE As Color = Color.Black
    Public Shared ReadOnly COLOURS_BUTTON_BORDER As Color = Color.Black
    Public Shared ReadOnly COLOURS_BUTTON_DISABLED_BORDER As Color = SystemColors.Control
    Public Shared ReadOnly COLOURS_BUTTON_MOUSEOVER As Color = Color.DodgerBlue
    Public Shared ReadOnly COLOURS_BUTTON_MOUSEDOWN As Color = Color.DeepSkyBlue
    Public Shared ReadOnly COLOURS_BUTTON_NEW_GAME_BACK As Color = Color.DodgerBlue
    Public Shared ReadOnly COLOURS_WINNER_TEXT As Color = Color.ForestGreen
    Public Shared ReadOnly COLOURS_LOSER_TEXT As Color = Color.Crimson
    Public Shared ReadOnly COLOURS_WORD_NORMAL As Color = Color.Black
    Public Shared ReadOnly COLOURS_WORD_HIGHLIGHTED As Color = Color.Firebrick


End Class