Option Explicit

Public Type tRaman
    Wavenumber As Double
    Intensity As Double
End Type

Public Function csvRaman(csvLine$) As tRaman

    Dim result As tRaman, _
        delimiter$, _
        position$, _
        length&

        delimiter = ","
        length = Len(csvLine)
        position = InStr(1, csvLine, delimiter)
        
        result.Wavenumber = CDbl(Left(csvLine, (position - 1)))
        result.Intensity = CDbl(Right(csvLine, (length - position)))
        
        csvRaman = result

End Function

Public Function csvFind$()

    Dim FDO As FileDialog, _
        SelectionChosen&
    
        Set FDO = Application.FileDialog(msoFileDialogFilePicker)
        SelectionChosen = -1
        
        With FDO
            .InitialFileName = "C:\"
            .Title = "Choose CSV"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Allowed File Extensions", "*.csv"
            
            If .Show = SelectionChosen Then
                csvFind = .SelectedItems(1)
            Else
            End If
            
        End With
    
        Set FDO = Nothing

End Function

Public Function csvParse(csvFilepath$, _
                         dataColumns&) As Double()

    Dim fileObject As Object, _
        textObject As Object, _
        dataRaman As tRaman, _
        dataLine$, _
        i&, _
        C#()
        
        Set fileObject = CreateObject("Scripting.FileSystemObject")
        'Set textObject = CreateObject("Scripting.TextStream")
        Set textObject = fileObject.OpenTextFile(csvFilepath)
        
        With textObject
        
            i = 0
            Do
                If .AtEndOfStream Then
                    Exit Do
                End If
            
                i = i + 1
                .SkipLine
            Loop
            
            .Close
            ReDim C(1 To i, 1 To dataColumns)
        
        End With
        
        Set textObject = fileObject.OpenTextFile(csvFilepath)
        
        With textObject
        
            i = 1
            Do
            
                If .AtEndOfStream Then
                    Exit Do
                End If
                
                dataLine = .ReadLine
                dataRaman = modText.csvRaman(dataLine)
                
                C(i, 1) = dataRaman.Wavenumber
                C(i, 2) = dataRaman.Intensity
                i = i + 1
            
            Loop
            .Close
        
        End With
        
        Set fileObject = Nothing
        Set textObject = Nothing
        
        csvParse = C
        
End Function

Public Function csvWrite(ByRef A#(), _
                         ByVal csvFilename$, _
                         Optional ByVal csvDirectory$ = "C:\Users\qp\Desktop\") As Boolean

    Dim FSO As Object, _
        txtFile As Object, _
        rawMaxRowA&, _
        rawMaxColA&, _
        concatString$, _
        delimiter$, _
        i&, j&
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set txtFile = FSO.CreateTextFile(csvDirectory & csvFilename)
        delimiter = ","
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
       
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                concatString = concatString & A(i, j) & delimiter
                
            Next j
            
            txtFile.Write concatString & vbCrLf
            concatString = ""
            
        Next i
        
        txtFile.Close
        csvWrite = True
        Set FSO = Nothing
        Set txtFile = Nothing

End Function

Function TempFolderExists() As Boolean

    Dim FSO As Object, _
        strFolder$
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        strFolder = "C:\temp"
        
        If FSO.FolderExists(strFolder) Then
            TempFolderExists = True
        Else
            FSO.CreateFolder (strFolder)
            TempFolderExists = True
        End If
        
        Set FSO = Nothing

End Function
