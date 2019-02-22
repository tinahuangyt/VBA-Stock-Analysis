Sub stockvol()
    Dim stockname as string
    Dim lastrow as long
    Dim summary_table_row as long
    Dim openprice as double 
    Dim closeprice as double
    Dim yearlychange as double 
    
    For Each ws in worksheets
        stocktotal = 0
        summary_table_row = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").value = "Ticker"
        ws.Range("L1").value = "Total stock volume"
        ws.Range("J1").value = "Yearly Change"
        ws.Range("K1").value = "Percent Change"
        
        for i = 2 to lastrow
            if ws.cells(i+1,1).value <> ws.cells(i,1).value then
                stockname = ws.cells(i,1).value
                stocktotal = stocktotal+ ws.cells(i,7).value
                'To calculate yearly change and percent change 
                closeprice = ws.cells(i,6).value
                yearlychange = closeprice - openprice
                percentchange = yearlychange/openprice  
                ws.Range("J"&summary_table_row).value = yearlychange
                'Conditional formatting red/green
                if yearlychange > 0 then 
                    ws.Range("J"&summary_table_row).Interior.ColorIndex=4
                    else 
                    ws.Range("J"&summary_table_row).Interior.ColorIndex=3
                    end if 
                ws.Range("I"&summary_table_row).value = stockname
                ws.Range("L"&summary_table_row).value = stocktotal
                ws.Range("K"&summary_table_row).value = percentchange
                summary_table_row=summary_table_row+1
                stocktotal=0
                openprice = 0
            else
                stocktotal=stocktotal+ws.cells(i,7).value
                if openprice= 0 then
                    openprice = ws.cells(i,3).value
                end if
            end if 
        next i 
        ' To change percent change cells to %
        for i = 2 to lastrow
            ws.cells(i,11).Style = "Percent"
        next i 
    next ws 
end sub 
