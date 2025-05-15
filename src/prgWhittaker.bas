Option Explicit

Sub WhittakerHenderson()

    Dim csvStatus As Boolean, _
        csvFilepath$, _
        csvColumn&, _
        csvMatrix#(), _
 _
        dataWave#(), _
        dataWave_dS#(), _
        dataInte#(), _
        dataInte_dS#(), _
        dataInte_primarySmooth#(), _
        dataInte_finalSmooth#(), _
        dataInte_searchSmooth#(), _
        dataInte_fD#(), _
        dataInte_peaks#(), _
        dataInte_valleys#(), _
        dataInte_valleys_advanced#(), _
        data_dS#(), _
        data_primarySmooth#(), _
        data_final#(), _
        i&
        
        csvColumn = 2
        csvFilepath = modText.csvFind
        
        If Len(csvFilepath) = 0 Then
            Exit Sub
        End If
        
        csvMatrix = modText.csvParse(csvFilepath, csvColumn)
        
        'Dim op#()
        'op = modOptimization.optPolyFit(csvMatrix, 6)
        'csvStatus = modText.csvWrite(op, "opt6.csv")
        'csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
        
'        Dim sav#(), pq#()
'        sav = modOptimization.optSavGol(csvMatrix, 51)
'        pq = modOptimization.optPeaks_SG(csvMatrix, sav)
'        csvStatus = modText.csvWrite(sav, "sav.csv")
'        csvStatus = modText.csvWrite(pq, "savPeaks.csv")
'        csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
'
        
        
        
        dataWave = modMatrix.matVec(csvMatrix, 1)
        dataWave_dS = modMath.mathDownSampling(dataWave, 1)
        dataInte = modMatrix.matVec(csvMatrix, 2)
        dataInte_dS = modMath.mathDownSampling(dataInte, 1)

        'dataInte_primarySmooth = modOptimization.optWhittaker(dataInte_dS, , 0, , 1)
        
        'dataInte_fD = modOptimization.optfD(dataWave_dS, dataInte_primarySmooth)
        
        'data_primarySmooth = modMatrix.matJoin(dataWave_dS, dataInte_primarySmooth)
        data_dS = modMatrix.matJoin(dataWave_dS, dataInte_dS)
        
        'dataInte_peaks = modOptimization.optPeaks(csvMatrix, data_dS, dataInte_fD)
        'dataInte_valleys = modOptimization.optValleys(csvMatrix, data_dS, dataInte_fD)
        'dataInte_valleys_advanced = modOptimization.optValleys_advanced(csvMatrix, data_dS, dataInte_fD, both_ways_and_ends)
        
        dataInte_finalSmooth = modOptimization.optWhittaker(dataInte_dS, 0.001, 50000#, 2, 7)
        'data_final = modMatrix.matJoin(dataWave_dS, dataInte_finalSmooth)
        
        'dataInte_searchSmooth = modOptimization.optWhittaker_search(csvMatrix, dataWave_dS, dataInte_dS, dataInte_valleys_advanced, dataInte_peaks, , , , 6, 10000000000#)
        
        
        
        'csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
        csvStatus = modText.csvWrite(data_dS, "dS.csv")
        'csvStatus = modText.csvWrite(data_primarySmooth, "pSmooth.csv")
        'csvStatus = modText.csvWrite(dataInte_peaks, "peaks.csv")
        'csvStatus = modText.csvWrite(dataInte_valleys, "valleys.csv")
        'csvStatus = modText.csvWrite(dataInte_valleys_advanced, "advanced_valleys.csv")
        csvStatus = modText.csvWrite(dataInte_finalSmooth, "final_smooth.csv")
        'csvStatus = modText.csvWrite(data_final, "final.csv")
        
        'csvStatus = modText.csvWrite(dataInte_searchSmooth, "search.csv")









End Sub
