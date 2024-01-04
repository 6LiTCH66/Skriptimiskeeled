inputFile = WScript.Arguments(0)
wordCountLimit = WScript.Arguments(1)

set textFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(inputFile, 1)
set wordFrequency = CreateObject("Scripting.Dictionary")

do until textFile.AtEndOfStream
    currentLine = LCase(textFile.ReadLine())
    wordsArray = Split(currentLine, " ")
    for each item In wordsArray
        item = Replace(item, ",", "")
        item = Replace(item, ";", "")
        item = Replace(item, ":", "")
        item = Replace(item, ")", "")
        item = Replace(item, "(", "")
        item = Replace(item, "!", "")
        item = Replace(item, "?", "")
        item = Replace(item, ".", "")
        item = Replace(item, """", "")
        item = Replace(item, "--", "")
        item = Replace(item, "&", "")
        item = Replace(item, "$", "")
        item = Replace(item, "  ", "")
        item = Replace(item, "_", "")

        select case item
            case "don't", "doesn't"
                item = "do not"
            case "can't"
                item = "cannot"
            case "an"
                item = "a"
            case "me", "mine", "my", "myself", "i'm", "'i"
                item = "i"
            case "him", "his", "himself"
                item = "he"
            case "her", "hers", "herself"
                item = "she"
            case "its", "itself"
                item = "it"
            case "whom", "whose"
                item = "who"
            case "one's", "oneself"
                item = "one"
            case "us", "our", "ours", "ourselves"
                item = "we"
            case "yours", "your", "yourself", "yourselves"
                item = "you"
            case "them", "their", "theirs", "themselves"
                item = "they"
            case "am", "is", "are", "was", "were", "being", "been"
                item = "be"
            case "has", "had", "having"
                item = "have"
            case "an'"
                item = "and"
        end Select

        if wordFrequency.Exists(item) then
            wordFrequency(item) = wordFrequency(item) + 1
        else
            wordFrequency.Add item, 1
        end if
    next
loop


wscript.echo "CHECKING THE ZIPF's LAW"
wscript.echo
wscript.echo "The first column is the number of corresponding words in the text and the second column is the number of words which should occur in the text according to the Zipf's law."
wscript.echo
wscript.echo "The most popular words in " + inputFile + "are:"
wscript.echo


For index = 0 To wordCountLimit
    maxWord = "" : maxCount = 0
    For Each term In wordFrequency
        If wordFrequency(term) > maxCount Then
            maxWord = term
            maxCount = wordFrequency(term)
        End If
    Next

    If maxCount <> 0 Then
        wordFrequency.Remove(maxWord)
        If Not (maxWord = vbLf Or maxWord = vbCr Or maxWord = "" Or maxWord = " ") Then
            zipfLawCount = maxCount
            WScript.Echo maxWord, vbTab, maxCount, vbTab, Round(zipfLawCount / index)
        End If
    End If
Next

wscript.echo
wscript.echo "The most popular still remaining short forms in" + inputFile + "are:"
wscript.echo


for index = 0 to wordCountLimit
    maxWord = "" : maxCount = 0
    for each term In wordFrequency
        if wordFrequency(term) > maxCount And InStr(term, "'") then
            maxWord = term
            maxCount = wordFrequency(term)
        end if
    next

    if maxCount <> 0 then
        wordFrequency.Remove(maxWord)
        if Not (maxWord = vbLf Or maxWord = vbCr Or maxWord = "" Or maxWord = " ") Then
            WScript.Echo maxWord, vbTab, maxCount
        end if
    end if
next