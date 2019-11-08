open System

#light
#I @"C:\Users\netha\Documents\FSharpTest\packages\Microsoft.Office.Interop.Word\lib\net20"
#r "Microsoft.Office.Interop.Word.dll"

module FTEST1 =

    open Microsoft.Office.Interop.Word
    open System.IO

    let comarg x = ref (box x)

    let printDocument (doc : Document) =
        printfn "Printing %s..." doc.Name

    let findAndReplace (doc : Document, findText : string, replaceWithText : string) =

        printfn "finding and replacing  %s..." doc.Name
        
        //options
        let matchCase = comarg false
        let matchWholeWord = comarg true
        let matchWildCards = comarg false
        let matchSoundsLike = comarg false
        let matchAllWordForms = comarg false
        let forward = comarg true
        let format = comarg false
        let matchKashida = comarg false
        let matchDiacritics = comarg false
        let matchAlefHamza = comarg false
        let matchControl = comarg false
        let read_only = comarg false
        let visible = comarg true
        let replace = comarg 2
        let wrap = comarg 1
        //execute find and replace

        let res = 
            doc.Content.Find.Execute(
                comarg findText, 
                matchCase, 
                matchWholeWord,
                matchWildCards, 
                matchSoundsLike, 
                matchAllWordForms, 
                forward, 
                wrap, 
                format, 
                comarg replaceWithText, 
                replace,
                matchKashida,
                matchDiacritics, 
                matchAlefHamza, 
                matchControl)

        printfn "Result of Execute is: %b" res

    let wordApp = new Microsoft.Office.Interop.Word.ApplicationClass(Visible = true)

    let openDocument fileName = 
        wordApp.Documents.Open(comarg fileName)

    // example useage
    let closeDocument (doc : Document) =
        printfn "Closing %s…" doc.Name
        // wdDoNotSaveChanges = 0 | wdPromptToSaveChanges = -2 | wdSaveChanges = -1
        doc.Close(SaveChanges = comarg -1)

    let findText = "test"
    let replaceText = "McTesty"

    let findandreplaceinfolders folder findText replaceText =
        Directory.GetFiles(folder, "*.docx")
        |> Array.iter (fun filePath ->
            let doc = openDocument filePath
            doc.Activate()
            // printDocument doc
            printfn "I work"
            findAndReplace(doc, findText, replaceText)
            closeDocument doc)


    let currentFolder = __SOURCE_DIRECTORY__

    printfn "Printing all files in [%s]..." currentFolder
    findandreplaceinfolders currentFolder findText replaceText
    
    wordApp.Quit()

printfn "Press any key…"
Console.ReadKey(true) |> ignore