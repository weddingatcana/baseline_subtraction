Option Explicit

Sub WLS()

    Dim csvStatus As Boolean, _
        csvFilepath$, _
        csvColumn&, _
        csvMatrix#(), _
 _
        dataWave#(), _
        dataWave_dS#(), _
        dataInte#(), _
        dataInte_dS#(), _
        data_dS#(), _
 _
        polyCoeffs#(), _
        polyFinal#()
        
        csvColumn = 2
        csvFilepath = modText.csvFind
        
        If Len(csvFilepath) = 0 Then
            Exit Sub
        End If
        
        csvMatrix = modText.csvParse(csvFilepath, csvColumn)
        
        dataWave = modMatrix.matVec(csvMatrix, 1)
        dataWave_dS = modMath.mathDownSampling(dataWave, 1)
        dataInte = modMatrix.matVec(csvMatrix, 2)
        dataInte_dS = modMath.mathDownSampling(dataInte, 1)
        data_dS = modMatrix.matJoin(dataWave_dS, dataInte_dS)
        
        polyCoeffs = modOptimization.optWeightedLeastSquares(dataWave_dS, dataInte_dS, 3, 0.001, 20)
        polyFinal = modOptimization.optPolyFit_seperate_coeff(dataWave_dS, polyCoeffs)

        csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
        csvStatus = modText.csvWrite(data_dS, "dS.csv")
        csvStatus = modText.csvWrite(polyFinal, "WLS.csv")
       
End Sub
