Sub MicroPrintToPDF1()

    '変数宣言
    Dim i As Integer
    
    '" Book1.xlsx "を開く
    Application.ScreenUpdating = False
        Workbooks.Open fileName:="C:\Users\***\OneDrive\Book1.xlsx"
          
    'ループ1
    For i = 1 To 1
    
    
    '*** 「sheet1」のページ設定：start ***
    
    With Worksheets("sheet1").PageSetup
               
    'シート「sheet1」をアクティブにする
    With Worksheets("sheet1").Select
    
    'B1～D17セルを印刷範囲として指定する
    Worksheets("sheet1").PageSetup.PrintArea = "B1:D17"
        
    '印刷向きを横方向に設定
    .Orientation = xlLandscape
                       
    'すべての列を1ページに印刷
    .FitToPagesWide = 1
    
    '*** 「sheet1」のページ設定：end ***
    
    
    'シートに中身がなければ印刷せずに次のページ設定に飛ぶ
    If Not IsError(Cells(4, 3)) Then
        If Not Cells(4, 3).HasFormula Then
            If Cells(4, 3) = "" Then
                Exit For
            End If
        End If
    End If

    End With
    
    '「 sheet1 」のシートを「 C:\PDF置き場 」にPDF出力(Microsoft Print to PDF形式)
    Worksheets("sheet1").PrintOut ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, PrToFileName:="C:\PDF置き場\Book1_sheet1" & ".pdf"
      
Next i

    'ループ2
    For i = 1 To 1
    
    
    '*** 「sheet2」のページ設定：start ***
    
    With Worksheets("sheet2").PageSetup
    
    'シート「sheet2」をアクティブにする
    Worksheets("sheet2").Select
               
    'B1～D17セルを印刷範囲として指定する
    Worksheets("sheet2").PageSetup.PrintArea = "B1:D16"
        
    '印刷向きを横方向に設定
    .Orientation = xlLandscape
                       
    'すべての列を1ページに印刷
    .FitToPagesWide = 1
    
    '*** 「sheet2」のページ設定：end ***
    
    
    'シートに中身がなければ印刷せずに次のページ設定に飛ぶ
    If Not IsError(Cells(4, 3)) Then
        If Not Cells(4, 3).HasFormula Then
            If Cells(4, 3) = "" Then
                Exit For
            End If
        End If
    End If

    End With
    
    '「sheet2 」のシートを「 C:\PDF置き場 」にPDF出力(microsoft print to pdf形式)
    Worksheets("sheet2").PrintOut ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, PrToFileName:="C:\PDF置き場\Book1_sheet2" & ".pdf"
               
Next i


    '*** 「sheet3」のページ設定：start ***
       
    'ループ3
    For i = 1 To 1
    
    With Worksheets("sheet3").PageSetup
    
    'シート「sheet3」をアクティブにする
    Worksheets("sheet3").Select
               
    ' B1～D17セルを印刷範囲として指定する
    Worksheets("sheet3").PageSetup.PrintArea = "B1:D17"
        
    ' 印刷向きを横方向に設定
    .Orientation = xlLandscape
                       
    ' すべての列を1ページに印刷
    .FitToPagesWide = 1
    
    End With
    
    ' *** 「sheet3」のページ設定：end ***
    
    
    ' シートに中身がなければ印刷せずに終了する
        If Not IsError(Cells(4, 3)) Then
            If Not Cells(4, 3).HasFormula Then
                If Cells(4, 3) = "" Then
                Exit For
            End If
        End If
    End If
    
    '「 sheet3 」のシートを「 C:\PDF置き場 」にPDF出力(Microsoft Print to PDF形式)
    Worksheets("sheet3").PrintPreview EnableChanges:=True
        Worksheets("sheet3").PrintOut ActivePrinter:="Microsoft Print to PDF", _
            PrintToFile:=True, PrToFileName:="C:\PDF置き場\Book1_sheet3" & ".pdf"
        
Next i
        
    '終了処理
    Application.DisplayAlerts = False
    Workbooks("知識・用語集.xlsx").Close
    Application.DisplayAlerts = True
    
    
End Sub
