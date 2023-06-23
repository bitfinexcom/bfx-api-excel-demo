Attribute VB_Name = "User_Defined_Functions"
Function BFX(Optional Symbol As String, Optional Field As String, Optional Timestamp As String, Optional Timeframe As String, Optional Period As String, Optional Second_Symbol As String, Optional Second_Timestamp As String)

'First, the Function declares the different variables requires to make an API call

Dim URL As String
Dim ApiData As String
Dim Pair As String
Dim Endpoint As String
Dim Query_Parameters As String
Dim Path_Parameters As String
Dim Data1 As Double
Dim Data2 As Double
Dim SymMem As String
Dim ErrorCheck As String

'The Bitfinex API prepends trading symbols with a 't' and funding symbols with an 'f'.
'Since there are nu funding symbols longer than four characters, the script can use this to distinguish trading and funding symbols in order to prepend them with the appropriate symbol.

If Len(Symbol) < 5 Then
    Pair = "f" + Symbol
Else
    Pair = "t" + Symbol
End If

'The script uses the field argument entered into the function to determine which endpoint and parameters should be used when buildling the URL.
'The script calls the BuildURL() function to combine the Endpoint, Path Parameters, and Query Parameters into a valid URL.
'The script call the SendRequest() function to send a call to the API and retrieve the response.
'The script calls the ReturnField function to determine which element of the API response should be returned to the Excel worksheet.

If Field = "Bid" Or Field = "Bid_Size" Or Field = "Ask" Or Field = "Ask_Size" Or Field = "Daily_Change" Or Field = "Daily_Change_Relative" Or Field = "Last_Price" Or Field = "Volume" And Timeframe = "" Or Field = "High" And Timeframe = "" Or Field = "Low" And Timeframe = "" Or Field = "FRR" Or Field = "Bid_Period" Or Field = "Ask_Period" Or Field = "FRR_AA" Then
    
    If Len(Symbol) > 4 And Field = "FRR" Or Len(Symbol) > 4 And Field = "Bid_Period" Or Len(Symbol) > 4 And Field = "Ask_Period" Or Len(Symbol) > 4 And Field = "FRR_AA" Then
        BFX = "Error, field not available for this SYMBOL"
    
    ElseIf Symbol <> "" Then
    
        Endpoint = "tickers"
        Query_Parameters = "symbols=" + Pair
        URL = BuildURL(Endpoint, , Query_Parameters)
    
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint, , , Symbol)
        Else
            BFX = ErrorCheck
        End If
    
    End If
    
ElseIf Field = "Last_Price_Hist" Or Field = "First_Price_Hist" Then
    
    If Symbol <> "" And Timestamp <> "" Then
    
        Endpoint = "trades"
        Path_Parameters = Pair + "/hist"
        
        If Field = "Last_Price_Hist" Then
            Query_Parameters = "end=" + Timestamp + "&limit=1"
        ElseIf Field = "First_Price_Hist" Then
            Query_Parameters = "start=" + Timestamp + "&limit=1" + "&sort=1"
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
    
    Else
        If Symbol = "" And Timestamp = "" Then
            BFX = "Error, SYMBOL and TIMESTAMP missing"
        ElseIf Symbol = "" Then
            BFX = "Error, SYMBOL missing"
        ElseIf Timestamp = "" Then
            BFX = "Error, TIMESTAMP missing"
        End If
        
    End If

ElseIf Field = "Mid_Price" Then

    If Symbol <> "" Then
    
        Endpoint = "book"
        Path_Parameters = Pair + "/P0"
        Query_Parameters = "len=1"
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
         
    Else
        BFX = "Error, SYMBOL missing"
    End If
    
    
ElseIf Field = "Platform_Status" Then
    
    Endpoint = "platform/status"
    Path_Parameters = ""
    Query_Parameters = ""
    URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
    
    ApiData = SendRequest(URL)
    
    BFX = ReturnField(ApiData, Field, Endpoint)

ElseIf Field = "Open" Or Field = "Close" Or Field = "High" And Timeframe <> "" Or Field = "Low" And Timeframe <> "" Or Field = "Volume" And Timeframe <> "" Then

    If Symbol <> "" And Len(Symbol) < 5 And Timeframe <> "" And Period = "" Then
        BFX = "Error, funding candles require a period to be specified"
    ElseIf Symbol <> "" And Timeframe <> "" Then
    
        Endpoint = "candles"
    
        If Period = "" Then
            Path_Parameters = "trade:" + Timeframe + ":" + Pair + "/hist"
        Else
            Path_Parameters = "trade:" + Timeframe + ":" + Pair + ":" + Period + "/hist"
        End If
        
        If Timestamp = "" Then
            Query_Parameters = "limit=1"
        Else
            Query_Parameters = "limit=1&end=" + Timestamp
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint, Timeframe)
        Else
            BFX = ErrorCheck
        End If
    
    ElseIf Symbol = "" Then
        BFX = "Errors, SYMBOL missing"
    ElseIf Timeframe = "" Then
        BFX = "Error, TIMEFRAME missing"
    End If
    
ElseIf Field = "Pos_Size_Long" Or Field = "Pos_Size_Short" Or Field = "Pos_Size_Long_Quote" Or Field = "Pos_Size_Short_Quote" Then

    If Symbol <> "" Then
    
        Endpoint = "stats1"
        
        If Field = "Pos_Size_Long" Or Field = "Pos_Size_Long_Quote" Then
            Path_Parameters = "pos.size:" + "1m:" + Pair + ":long" + "/hist"
        ElseIf Field = "Pos_Size_Short" Or Field = "Pos_Size_Short_Quote" Then
            Path_Parameters = "pos.size:" + "1m:" + Pair + ":short" + "/hist"
        End If
        
        If Timestamp = "" Then
            Query_Parameters = "limit=1"
        Else
            Query_Parameters = "limit=1&end=" + Timestamp
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            
            If Field = "Pos_Size_Long" Or Field = "Pos_Size_Short" Then
            
                BFX = ReturnField(ApiData, Field, Endpoint)
        
            ElseIf Field = "Pos_Size_Long_Quote" Or Field = "Pos_Size_Short_Quote" Then
            
                Data1 = ReturnField(ApiData, "Pos_Size_Long", Endpoint)
                
                If Timestamp = "" Then
                    Data2 = BFX(Symbol, "Bid")
                Else
                    Data2 = BFX(Symbol, "Last_Price_Hist", Timestamp)
                End If
                
                If CheckErrorReturned(Data2) = "Error" Then
                    BFX = Data2
                Else
                    BFX = Data1 * Data2
                End If
            
            End If
            
        Else
            BFX = ErrorCheck
        End If
    
    ElseIf Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    End If

ElseIf Field = "Funding_Size" Or Field = "Credits_Size" Or Field = "Credits_Size_Sym" Then

    If Symbol <> "" Then
        
        Endpoint = "stats1"
        
        If Field = "Funding_Size" Then
            Path_Parameters = "funding.size" + ":" + "1m" + ":" + Pair + "/hist"
        ElseIf Field = "Credits_Size_Sym" Then
            Path_Parameters = "credits.size.sym" + ":" + "1m" + ":" + Pair + ":" + "t" + Second_Symbol + "/hist"
        ElseIf Field = "Credits_Size" Then
            Path_Parameters = "credits.size" + ":" + "1m" + ":" + Pair + "/hist"
        End If
        
        If Timestamp = "" Then
            Query_Parameters = "limit=1"
        Else
            Query_Parameters = "limit=1&end=" + Timestamp
        End If
        
        If Field = "credits.size.sym" And Second_Symbol = "" Then
            BFX = "Error, SECOND_SYMBOL Missing"
        ElseIf Field = "credits.size.sym" And Len(Symbol) > 4 Then
            BFX = "Error, SYMBOL needs to be a funding currency"
        ElseIf Field = "Credits_Size_Sym" And Len(Second_Symbol) < 4 Then
            BFX = "Error, SECOND_SYMBOL needs to be a trading pair"
        Else
        
            URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
            ApiData = SendRequest(URL)
            
            ErrorCheck = CheckErrors(ApiData)
        
            If ErrorCheck = "Pass" Then
                BFX = ReturnField(ApiData, Field, Endpoint)
            Else
                BFX = ErrorCheck
            End If
        
        End If
    
    ElseIf Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    End If

ElseIf Field = "Vol.1d" Or Field = "Vol.7d" Or Field = "Vol.30d" Or Field = "Vwap" Then
    
    Endpoint = "stats1"
    
    If Field = "Vol.1d" Then
        Path_Parameters = "vol.1d:30m:BFX/hist"
    ElseIf Field = "Vol.7d" Then
        Path_Parameters = "vol.7d:30m:BFX/hist"
    ElseIf Field = "Vol.30d" Then
        Path_Parameters = "vol.30d:30m:BFX/hist"
    ElseIf Field = "Vwap" Then
        Path_Parameters = "vwap:1d" + ":" + Pair + "/hist"
    End If
    
    If Timestamp = "" Then
        Query_Parameters = "limit=1"
    ElseIf Timestamp <> "" Then
        Query_Parameters = "limit=1&end=" + Timestamp
    End If
    
    If Field = "Vwap" And Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    Else
    
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
    
    End If
    
ElseIf Field = "Deriv_Mid_Price" Or Field = "Spot_Price_Underlying" Or Field = "Insurance_Fund_Balance" Or Field = "Next_Funding_TS" Or Field = "Next_Funding_Accrued" Or Field = "Next_Funding_Step" Or Field = "Current_Funding" Or Field = "Mark_Price" Or Field = "Open_Interest" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    Else
    
        Endpoint = "status"
        
        If Timestamp = "" Then
            Path_Parameters = "deriv"
            Query_Parameters = "keys=" + Pair + "&limit=1"
        ElseIf Timestamp <> "" Then
            Path_Parameters = "deriv/" + Pair + "/hist"
            Query_Parameters = "limit=1&end=" + Timestamp
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint, , Timestamp)
        Else
            BFX = ErrorCheck
        End If
        
    End If
    
ElseIf Field = "Custom_Volume" Or Field = "Custom_High" Or Field = "Custom_Low" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Timeframe = "" Then
        BFX = "Error, TIMEFRAME missing"
    ElseIf Len(Symbol) < 5 And Period = "" Then
        BFX = "Error, PERIOD is required for funding currency data"
    ElseIf Second_Timestamp <> "" And Second_Timestamp < Timestamp Then
        BFX = "Error, TIMESTAMP cannot be greater than SECOND_TIMESTAMP"
    Else

    Endpoint = "candles"
    
        If Len(Symbol) < 5 Then
            Path_Parameters = "trade" + ":" + Timeframe + ":" + Pair + ":" + Period + "/hist"
        Else
            Path_Parameters = "trade" + ":" + Timeframe + ":" + Pair + "/hist"
        End If
        
        Path_Parameters = "trade:" + Timeframe + ":" + Pair + "/hist"
        Query_Parameters = "limit=10000&start=" + Timestamp + "&end=" + Second_Timestamp
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
    
    End If
        
ElseIf Field = "1Y_Avg_Volume" Or Field = "26W_Avg_Volume" Or Field = "13W_Avg_Volume" Or Field = "4W_Avg_Volume" Or Field = "1Y_Volume" Or Field = "26W_Volume" Or Field = "13W_Volume" Or Field = "4W_Volume" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Len(Symbol) < 5 And Period = "" Then
        BFX = "Error, PERIOD is required for funding currency data"
    Else
        
        Endpoint = "candles"
        
        If Len(Symbol) < 5 Then
            Path_Parameters = "trade:1D" + ":" + Pair + ":" + Period + "/hist"
        Else
            Path_Parameters = "trade:1D" + ":" + Pair + "/hist"
        End If
        
        If Field = "1Y_Avg_Volume" Or Field = "1Y_Volume" Then
            Query_Parameters = "limit=365"
        ElseIf Field = "26W_Avg_Volume" Or Field = "26W_Volume" Then
            Query_Parameters = "limit=182"
        ElseIf Field = "13W_Avg_Volume" Or Field = "13W_Volume" Then
            Query_Parameters = "limit=91"
        ElseIf Field = "4W_Avg_Volume" Or Field = "4W_Volume" Then
            Query_Parameters = "limit=28"
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
    
    End If
   
ElseIf Field = "High1Y" Or Field = "Low1Y" Or Field = "High26W" Or Field = "Low26W" Or Field = "High13W" Or Field = "Low13W" Or Field = "High4W" Or Field = "Low4W" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Len(Symbol) < 5 And Period = "" Then
        BFX = "Error, PERIOD is required for funding currency data"
    Else
    
        Endpoint = "candles"
        
        If Len(Symbol) < 5 Then
            Path_Parameters = "trade:1D" + ":" + Pair + ":" + Period + "/hist"
        Else
            Path_Parameters = "trade:1D" + ":" + Pair + "/hist"
        End If
        
        If Field = "High1Y" Or Field = "Low1Y" Then
            Query_Parameters = "limit=365"
        ElseIf Field = "High26W" Or Field = "Low26W" Then
            Query_Parameters = "limit=182"
        ElseIf Field = "High13W" Or Field = "Low13W" Then
            Query_Parameters = "limit=91"
        ElseIf Field = "High4W" Or Field = "Low4W" Then
            Query_Parameters = "limit=28"
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
            BFX = ReturnField(ApiData, Field, Endpoint)
        Else
            BFX = ErrorCheck
        End If
    
    End If
    
ElseIf Field = "1Y_Change" Or Field = "26W_Change" Or Field = "13W_Change" Or Field = "4W_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    Else
    
        Endpoint = "candles"
        Path_Parameters = "trade:1D" + ":" + Pair + "/hist"
        
        If Field = "1Y_Change" Then
            Query_Parameters = "limit=365"
        ElseIf Field = "26W_Change" Then
            Query_Parameters = "limit=182"
        ElseIf Field = "13W_Change" Then
            Query_Parameters = "limit=91"
        ElseIf Field = "4W_Change" Then
            Query_Parameters = "limit=28"
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
        
            If Field = "1Y_Change" Then
                Data1 = ReturnField(ApiData, "1YOpen", Endpoint)
            ElseIf Field = "26W_Change" Then
                Data1 = ReturnField(ApiData, "26WOpen", Endpoint)
            ElseIf Field = "13W_Change" Then
                Data1 = ReturnField(ApiData, "13WOpen", Endpoint)
            ElseIf Field = "4W_Change" Then
                Data1 = ReturnField(ApiData, "4WOpen", Endpoint)
            End If
            
            Endpoint = "book"
            Path_Parameters = Pair + "/P0"
            Query_Parameters = "len=1"
            URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
            ApiData = SendRequest(URL)
            
            ErrorCheck = CheckErrors(ApiData)
        
            If ErrorCheck = "Pass" Then
                Data2 = ReturnField(ApiData, "Mid_Price", Endpoint)
                BFX = Data2 - Data1
            Else
            BFX = ErrorCheck
            
            End If
            
        Else
            BFX = ErrorCheck
            
        End If
    
    End If
    
ElseIf Field = "1Y_Change_Relative" Or Field = "26W_Change_Relative" Or Field = "13W_Change_Relative" Or Field = "4W_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    Else
    
        Endpoint = "candles"
        Path_Parameters = "trade:1D" + ":" + Pair + "/hist"
        
        If Field = "1Y_Change_Relative" Then
            Query_Parameters = "limit=365"
        ElseIf Field = "26W_Change_Relative" Then
            Query_Parameters = "limit=182"
        ElseIf Field = "13W_Change_Relative" Then
            Query_Parameters = "limit=91"
        ElseIf Field = "4W_Change_Relative" Then
            Query_Parameters = "limit=28"
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
    
        If ErrorCheck = "Pass" Then
            
            If Field = "1Y_Change_Relative" Then
                Data1 = ReturnField(ApiData, "1YOpen", Endpoint)
            ElseIf Field = "26W_Change_Relative" Then
                Data1 = ReturnField(ApiData, "26WOpen", Endpoint)
            ElseIf Field = "13W_Change_Relative" Then
                Data1 = ReturnField(ApiData, "13WOpen", Endpoint)
            ElseIf Field = "4W_Change_Relative" Then
                Data1 = ReturnField(ApiData, "4WOpen", Endpoint)
            End If
            
            Endpoint = "book"
            Path_Parameters = Pair + "/P0"
            Query_Parameters = "len=1"
            URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
            ApiData = SendRequest(URL)
            
            ErrorCheck = CheckErrors(ApiData)
            
            If ErrorCheck = "Pass" Then
            
                Data2 = ReturnField(ApiData, "Mid_Price", Endpoint)
                
                BFX = (Data2 - Data1) / Data1 * 100
            
            Else
                
                BFX = ErrorCheck
            
            End If
        
        Else
        
        BFX = ErrorCheck
        
        End If
    
    End If

ElseIf Field = "Custom_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Timestamp <> "" And Timestamp > Second_Timestamp Then
        BFX = "Error, TIMESTAMP cannot be greater than SECOND_TIMESTAMP"
    Else
    
        Endpoint = "trades"
        Path_Parameters = Pair + "/hist"
        Query_Parameters = "end=" + Timestamp + "&limit=1"
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
        
            Data1 = ReturnField(ApiData, "Last_Price_Hist")
            
            If Second_Timestamp = "" Then
            
                Endpoint = "book"
                Path_Parameters = Pair + "/P0"
                Query_Parameters = "len=1"
                URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
                ApiData = SendRequest(URL)
                
                ErrorCheck = CheckErrors(ApiData)
                
                If ErrorCheck = "Pass" Then
            
                    Data2 = ReturnField(ApiData, "Mid_Price", Endpoint)
                
                Else
                    
                    Date2 = ErrorCheck
                
                End If
            
            Else
                
                Endpoint = "trades"
                Path_Parameters = Pair + "/hist"
                Query_Parameters = "end=" + Second_Timestamp + "&limit=1"
                URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
                ApiData = SendRequest(URL)
                
                ErrorCheck = CheckErrors(ApiData)
                
                If ErrorCheck = "Pass" Then
                
                    Data2 = ReturnField(ApiData, "Last_Price_Hist")
                
                Else
                
                    Data2 = ErrorCheck
                
                End If
            
            End If
        
            If CheckErrorReturned(Data2) = False Then

                BFX = Data2 - Data1
            
            Else
                
                BFX = Data2
            
            End If
        
        Else
            
            BFX = ErrorCheck
        
        End If
    
    End If
    
ElseIf Field = "Custom_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Timestamp <> "" And Timestamp > Second_Timestamp Then
        BFX = "Error, TIMESTAMP cannot be greater than SECOND_TIMESTAMP"
    Else
    
        Endpoint = "trades"
        Path_Parameters = Pair + "/hist"
        Query_Parameters = "end=" + Timestamp + "&limit=1"
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
        
            Data1 = ReturnField(ApiData, "Last_Price_Hist", Endpoint)
            
            If Second_Timestamp = "" Then
                
                Endpoint = "book"
                Path_Parameters = Pair + "/P0"
                Query_Parameters = "len=1"
                URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
                
                ApiData = SendRequest(URL)
                
                ErrorCheck = CheckErrors(ApiData)
        
                If ErrorCheck = "Pass" Then
                
                    Data2 = ReturnField(ApiData, "Mid_Price", Endpoint)
                    
                Else
                
                    Data2 = ErrorCheck
                    
                End If
            
            Else
                
                Endpoint = "trades"
                Path_Parameters = Pair + "/hist"
                Query_Parameters = "end=" + Second_Timestamp + "&limit=1"
                URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
                ApiData = SendRequest(URL)
                
                ErrorCheck = CheckErrors(ApiData)
                
                If ErrorCheck = "Pass" Then
            
                    Data2 = ReturnField(ApiData, "Last_Price_Hist")
                
                Else
                
                    Data2 = ErrorCheck
                
                End If
            
            End If
            
            If CheckErrorReturned(Data2) = False Then
            
                BFX = (Data2 - Data1) / Data1 * 100
            
            Else
            
                BFX = Data2
            
            End If
        
        Else
            BFX = Data1
        End If
    
    End If

ElseIf Field = "Percent_Long" Or Field = "Percent_Short" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Len(Symbol) < 5 Then
        BFX = "Error, SYMBOL must be a trading pair"
    Else
    
        Endpoint = "stats1"
        Path_Parameters = "pos.size:" + "1m:" + Pair + ":long" + "/hist"
        
        If Timestamp = "" Then
            Query_Parameters = "limit=1"
        Else
            Query_Parameters = "limit=1&end=" + Timestamp
        End If
        
        URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
        
        ApiData = SendRequest(URL)
        
        ErrorCheck = CheckErrors(ApiData)
        
        If ErrorCheck = "Pass" Then
        
            Data1 = ReturnField(ApiData, "Pos_Size_Long", Endpoint)
            
            
            Path_Parameters = "pos.size:" + "1m:" + Pair + ":short" + "/hist"
            
            If Timestamp = "" Then
                Query_Parameters = "limit=1"
            Else
                Query_Parameters = "limit=1&end=" + Timestamp
            End If
            
            URL = BuildURL(Endpoint, Path_Parameters, Query_Parameters)
            
            ApiData = SendRequest(URL)
            
            ErrorCheck = CheckErrors(ApiData)
            
            If ErrorCheck = "Pass" Then
            
                Data2 = ReturnField(ApiData, "Pos_Size_Short", Endpoint)
                    
                If Field = "Percent_Long" Then
                    BFX = (Data1 / (Data1 + Data2)) * 100
                ElseIf Field = "Percent_Short" Then
                    BFX = (Data2 / (Data1 + Data2)) * 100
                End If
            Else
            
            BFX = ErrorCheck
            
            End If
        
        Else
            
            BFX = ErrorCheck
        
        End If
    
    End If
    
ElseIf Field = "Longs_Minus_Borrowed" Or Field = "Shorts_Minus_Borrowed" Or Field = "Longs_Borrowed_Ratio" Or Field = "Shorts_Borrowed_Ratio" Then
    
    SymMem = Symbol
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Second_Symbol = "" Then
        BFX = "Error, SECOND_SYMBOL missing"
    ElseIf Len(Symbol) > 4 Then
        BFX = "Error, SYMBOL should be a funding currency"
    ElseIf Len(Second_Symbol) < 5 Then
        BFX = "Error, SECOND_SYMBOL should be a trading pair"
    Else
    
        If Field = "Longs_Minus_Borrowed" Or Field = "Longs_Borrowed_Ratio" Then
            If Timestamp = "" Then
                Data1 = BFX(Second_Symbol, "Pos_Size_Long_Quote")
            Else
                Data1 = BFX(Second_Symbol, "Pos_Size_Long_Quote", Timestamp)
            End If
        ElseIf Field = "Shorts_Minus_Borrowed" Or Field = "Shorts_Borrowed_Ratio" Then
            If Timestamp = "" Then
                Data1 = BFX(Second_Symbol, "Pos_Size_Short")
            Else
                Data1 = BFX(Second_Symbol, "Pos_Size_Short", Timestamp)
            End If
        End If
        
        If CheckErrorReturned(Data1) = False Then
        
            If Timestamp = "" Then
                Data2 = BFX(SymMem, "Credits_Size_Sym", , , , Second_Symbol)
            Else
                Data2 = BFX(SymMem, "Credits_Size_Sym", Timestamp, , , Second_Symbol)
            End If
            
            If CheckErrorReturned(Data2) = False Then
            
                If Field = "Shorts_Minus_Borrowed" Or Field = "Longs_Minus_Borrowed" Then
                    BFX = Data1 - Data2
                ElseIf Field = "Longs_Borrowed_Ratio" Or Field = "Shorts_Borrowed_Ratio" Then
                    BFX = Data1 / Data2
                End If
            
            Else
            
                BFX = Data2
                
            End If
        
        Else
        
            BFX = Data1
        
        End If
    
    End If

ElseIf Field = "Long_Size_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Pos_Size_Long", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
    
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Long", Second_Timestamp)
                    If CheckErrorReturned(Data2) = False Then
                        
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    
                    End If
                        
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                If CheckErrorReturned(Data2) = False Then
                    Data2 = BFX(Symbol, "Pos_Size_Long")
                    BFX = Data2 - Data1
                Else
                BFX = Data2
                
                End If
                
            End If
        
        Else
        
        BFX = Data1
        
        End If
    
    End If
    
ElseIf Field = "Long_Size_Change_Quote" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Pos_Size_Long_Quote", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Long_Quote", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Long_Quote")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
            End If
        
        Else
        
        BFX = Data1
        
        End If
    
    End If
    
ElseIf Field = "Long_Size_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
        
        Data1 = BFX(Symbol, "Pos_Size_Long", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Long", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Long")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
            
        Else
        
        BFX = Data1
        
        End If
        
    End If

ElseIf Field = "Long_Size_Quote_Change_Relative" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else

        Data1 = BFX(Symbol, "Pos_Size_Long_Quote", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Long_Quote", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Long_Quote")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
        
        BFX = Data1
        
        End If
    
    End If
    
ElseIf Field = "Short_Size_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Pos_Size_Short", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Short", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Short")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        
        End If
        
    End If
    
ElseIf Field = "Short_Size_Change_Quote" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Pos_Size_Short_Quote", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
            
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Short_Quote", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Short_Quote")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        End If
    
    End If
    
ElseIf Field = "Short_Size_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
    Data1 = BFX(Symbol, "Pos_Size_Short", Timestamp)
    
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Short", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Short")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                End If
                
            End If
            
        Else
            BFX = Data1
        End If
        
    End If

ElseIf Field = "Short_Size_Quote_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Pos_Size_Short_Quote", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Pos_Size_Short_Quote", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Pos_Size_Short_Quote")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
        
        BFX = Data1
        
        End If
    
    End If
    
ElseIf Field = "Funding_Size_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
    Data1 = BFX(Symbol, "Funding_Size", Timestamp)
    
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Funding_Size", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Funding_Size")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
                
            End If
            
        Else
            BFX = Data1
            
        End If
    
    End If
    
ElseIf Field = "Funding_Size_Change_Relative" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else

    Data1 = BFX(Symbol, "Funding_Size", Timestamp)
    
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Funding_Size", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Funding_Size")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        End If
    
    End If
    
ElseIf Field = "Credits_Size_Change" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else

        Data1 = BFX(Symbol, "Credits_Size", Timestamp)
        
        If CheckErrorReturned(Data) = False Then
    
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Credits_Size", Second_Timestamp)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Credits_Size")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        End If
    
    End If
     
ElseIf Field = "Credits_Size_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    Else
    
        Data1 = BFX(Symbol, "Credits_Size", Timestamp)
        
        If CheckErrorReturned(Data2) = False Then
    
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Credits_Size", Second_Timestamp)
                        
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Credits_Size")
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        End If
    
    End If
    
ElseIf Field = "Credits_Size_Sym_Change" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Symbol = "" Then
        BFX = "Error, SECOND_SYMBOL missing"
    ElseIf Len(Symbol) > 4 Then
        BFX = "Error, SYMBOL needs to be a funding currency"
    ElseIf Len(Second_Symbol) < 5 Then
        BFX = "Error, SECOND_SYMBOL needs to be a trading pair"
    Else
    
        Data1 = BFX(Symbol, "Credits_Size_Sym", Timestamp, , , Second_Symbol)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Credits_Size_Sym", Second_Timestamp, , , Second_Symbol)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = Data2 - Data1
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Credits_Size_Sym", , , , Second_Symbol)
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
                
            End If
        
        Else
            BFX = Data1
        End If
    
    End If
    
ElseIf Field = "Credits_Size_Sym_Change_Relative" Then
    
    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Symbol = "" Then
        BFX = "Error, SECOND_SYMBOL missing"
    ElseIf Len(Symbol) > 4 Then
        BFX = "Error, SYMBOL needs to be a funding currency"
    ElseIf Len(Second_Symbol) < 5 Then
        BFX = "Error, SECOND_SYMBOL needs to be a trading pair"
    Else
    
        Data1 = BFX(Symbol, "Credits_Size_Sym", Timestamp, , , Second_Symbol)
        
        If CheckErrorReturned(Data1) = False Then
        
            If Second_Timestamp <> "" Then
                If Second_Timestamp > Timestamp Then
                    Data2 = BFX(Symbol, "Credits_Size_Sym", Second_Timestamp, , , Second_Symbol)
                    
                    If CheckErrorReturned(Data2) = False Then
                        BFX = ((Data2 - Data1) / Data1) * 100
                    Else
                        BFX = Data2
                    End If
                    
                Else
                    BFX = "Error, Timestamp must be before Second_Timestamp"
                End If
            ElseIf Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Credits_Size_Sym", , , , Second_Symbol)
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
                
            End If
            
            BFX = ((Data2 - Data1) / Data1) * 100
        
        Else
            BFX = Data1
        End If
    
    End If
    
ElseIf Field = "Open_Interest_Change" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Timestamp <> "" And Timestamp > Second_Timestamp Then
        BFX = "Error, TIMESTAMP must be before SECOND_TIMESTAMP"
    Else
    
        Data1 = BFX(Symbol, "Open_Interest", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
            
            If Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Open_Interest")
                
                If CheckErrorReturned(Data2) = False Then
                    
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
            
            ElseIf Second_Timestamp <> "" Then
            
                Data2 = BFX(Symbol, "Open_Interest", Second_Timestamp)
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = Data2 - Data1
                Else
                    BFX = Data2
                End If
            
            End If
        
        Else
            BFX = Data1
            
        End If
    
    End If

ElseIf Field = "Open_Interest_Change_Relative" Then

    If Symbol = "" Then
        BFX = "Error, SYMBOL missing"
    ElseIf Timestamp = "" Then
        BFX = "Error, TIMESTAMP missing"
    ElseIf Second_Timestamp <> "" And Timestamp > Second_Timestamp Then
        BFX = "Error, TIMESTAMP must be before SECOND_TIMESTAMP"
    Else
    
        Data1 = BFX(Symbol, "Open_Interest", Timestamp)
        
        If CheckErrorReturned(Data1) = False Then
            
            If Second_Timestamp = "" Then
                Data2 = BFX(Symbol, "Open_Interest")
                
                If CheckErrorReturned(Data2) = False Then
                    
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
            
            ElseIf Second_Timestamp <> "" Then
            
                Data2 = BFX(Symbol, "Open_Interest", Second_Timestamp)
                
                If CheckErrorReturned(Data2) = False Then
                    BFX = ((Data2 - Data1) / Data1) * 100
                Else
                    BFX = Data2
                End If
            
            End If
        
        Else
            BFX = Data1
            
        End If
    
    End If
    
Else
    BFX = "Error, invalid field"
    
End If


End Function

Function BuildURL(Endpoint As String, Optional Path_Parameters As String, Optional Query_Parameters As String)

'Builds the URL to combine the BaseURL, endpoint, and parameters

    Dim BaseURL As String
    BaseURL = "https://api-pub.bitfinex.com/v2/"
    
    BuildURL = BaseURL + Endpoint + "/" + Path_Parameters + "?" + Query_Parameters

End Function

Function SendRequest(URL As String)

'Sends the request and receives the response
'WinHttpRequest is required for this function to work. This will need to be enabled in Microsoft Visual Basic for Applications under Tools > References

Dim req As WinHttpRequest
Dim Response As String

Set req = New WinHttpRequest
req.Open "GET", URL
req.Send

Response = req.ResponseText

SendRequest = Response

End Function

Function ReturnField(ApiData As String, Field As String, Optional Endpoint As String, Optional Timeframe As String, Optional Timestamp As String, Optional Symbol As String)

'Converts the response to a JSON object and then selects the appropriate index based upon the requested field, associated endpoint, and parameters

Dim ResponseObject As Object
Dim AggregateVolume As Double
Dim High As Double
Dim Low As Double

Set ResponseObject = JsonConverter.ParseJson(ApiData)

If Endpoint = "tickers" Then

    If Len(Symbol) < 5 Then
        
        If Field = "FRR" Then
            ReturnField = ResponseObject(1)(2)

        ElseIf Field = "Bid" Then
            ReturnField = ResponseObject(1)(3)
        
        ElseIf Field = "Bid_Period" Then
            ReturnField = ResponseObject(1)(4)

        ElseIf Field = "Bid_Size" Then
            ReturnField = ResponseObject(1)(5)
    
        ElseIf Field = "Ask" Then
            ReturnField = ResponseObject(1)(6)
            
        ElseIf Field = "Ask_Period" Then
            ReturnField = ResponseObject(1)(7)
    
        ElseIf Field = "Ask_Size" Then
            ReturnField = ResponseObject(1)(8)
    
        ElseIf Field = "Daily_Change" Then
            ReturnField = ResponseObject(1)(9)
    
        ElseIf Field = "Daily_Change_Relative" Then
            ReturnField = ResponseObject(1)(10)
    
        ElseIf Field = "Last_Price" Then
            ReturnField = ResponseObject(1)(11)
    
        ElseIf Field = "Volume" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(12)
    
        ElseIf Field = "High" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(13)
    
        ElseIf Field = "Low" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(14)
        
        ElseIf Field = "FRR_AA" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(17)
        
        End If
    
    Else
    
        If Field = "Bid" Then
            ReturnField = ResponseObject(1)(2)

        ElseIf Field = "Bid_Size" Then
            ReturnField = ResponseObject(1)(3)
    
        ElseIf Field = "Ask" Then
            ReturnField = ResponseObject(1)(4)
    
        ElseIf Field = "Ask_Size" Then
            ReturnField = ResponseObject(1)(5)
    
        ElseIf Field = "Daily_Change" Then
            ReturnField = ResponseObject(1)(6)
    
        ElseIf Field = "Daily_Change_Relative" Then
            ReturnField = ResponseObject(1)(7)
    
        ElseIf Field = "Last_Price" Then
            ReturnField = ResponseObject(1)(8)
    
        ElseIf Field = "Volume" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(9)
    
        ElseIf Field = "High" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(10)
    
        ElseIf Field = "Low" And Timeframe = "" Then
            ReturnField = ResponseObject(1)(11)
        
        End If
    
    End If

    

ElseIf Endpoint = "candles" Then
    
    If Field = "Open" Then
        ReturnField = ResponseObject(1)(2)
    
    ElseIf Field = "Close" Then
        ReturnField = ResponseObject(1)(3)

    ElseIf Field = "High" Then
        ReturnField = ResponseObject(1)(4)

    ElseIf Field = "Low" Then
        ReturnField = ResponseObject(1)(5)

    ElseIf Field = "Volume" Then
        ReturnField = ResponseObject(1)(6)
    ElseIf Field = "Custom_Volume" Or Field = "Old_Volume" Then

    If ResponseObject.Count < 10000 Then
        AggregateVolume = 0
        For Each Entry In ResponseObject
            AggregateVolume = AggregateVolume + Entry(6)
        Next
        ReturnField = AggregateVolume
    Else
        ReturnField = "Candle limit exceeded. Increase candle size or reduce the spread between the first and second timestamp"
    End If
    
    ElseIf Field = "Custom_High" Then
    
    If ResponseObject.Count < 10000 Then
        High = 0
        For Each Entry In ResponseObject
            If Entry(4) > High Then
                High = Entry(4)
            End If
        Next
        ReturnField = High
    Else
        ReturnField = "Candle limit exceeded. Increase candle size or reduce the spread between the first and second timestamp"
    End If
    
    ElseIf Field = "Custom_Low" Then
    
    If ResponseObject.Count < 10000 Then
        Low = 9.22337203685478E+18
        For Each Entry In ResponseObject
            If Entry(5) < Low Then
                Low = Entry(5)
            End If
        Next
        ReturnField = Low
    Else
        ReturnField = "Candle limit exceeded. Increase candle size or reduce the spread between the first and second timestamp"
    End If
    
    ElseIf Field = "1Y_Avg_Volume" Or Field = "26W_Avg_Volume" Or Field = "13W_Avg_Volume" Or Field = "4W_Avg_Volume" Or Field = "1Y_Volume" Or Field = "26W_Volume" Or Field = "13W_Volume" Or Field = "4W_Volume" Then

        AggregateVolume = 0
        For Each Entry In ResponseObject
            AggregateVolume = AggregateVolume + Entry(6)
        Next
        
        If Field = "1Y_Volume" Or Field = "26W_Volume" Or Field = "13W_Volume" Or Field = "4W_Volume" Then
            ReturnField = AggregateVolume
        ElseIf Field = "1Y_Avg_Volume" Then
            ReturnField = AggregateVolume / 365
        ElseIf Field = "26W_Avg_Volume" Then
            ReturnField = AggregateVolume / 182
        ElseIf Field = "13W_Avg_Volume" Then
            ReturnField = AggregateVolume / 91
        ElseIf Field = "4W_Avg_Volume" Then
            ReturnField = AggregateVolume / 28
        End If
    
    ElseIf Field = "High1Y" Or Field = "High26W" Or Field = "High13W" Or Field = "High4W" Then

        High = 0
        For Each Entry In ResponseObject
            If Entry(4) > High Then
                High = Entry(4)
            End If
        Next
    ReturnField = High
    
    ElseIf Field = "Low1Y" Or Field = "Low26W" Or Field = "Low13W" Or Field = "Low4W" Then

        Low = 9.22337203685478E+18
        For Each Entry In ResponseObject
            If Entry(5) < Low Then
                Low = Entry(5)
            End If
        Next
        ReturnField = Low
    
    ElseIf Field = "1YOpen" Then
        ReturnField = ResponseObject(365)(2)
    
    ElseIf Field = "26WOpen" Then
        ReturnField = ResponseObject(182)(2)
    
    ElseIf Field = "13WOpen" Then
        ReturnField = ResponseObject(91)(2)
    
    ElseIf Field = "4WOpen" Then
        ReturnField = ResponseObject(28)(2)
        
    End If

ElseIf Endpoint = "status" Then

    If Field = "Deriv_Mid_Price" Then
    
        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(3)
        Else
            ReturnField = ResponseObject(1)(4)
        End If
    
    ElseIf Field = "Spot_Price_Underlying" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(4)
        Else
            ReturnField = ResponseObject(1)(5)
        End If

    ElseIf Field = "Insurance_Fund_Balance" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(6)
        Else
            ReturnField = ResponseObject(1)(7)
        End If
    
    ElseIf Field = "Next_Funding_TS" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(8)
        Else
            ReturnField = ResponseObject(1)(9)
        End If

    ElseIf Field = "Next_Funding_Accrued" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(9)
        Else
            ReturnField = ResponseObject(1)(10)
        End If

    ElseIf Field = "Next_Funding_Step" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(10)
        Else
            ReturnField = ResponseObject(1)(11)
        End If
    
    ElseIf Field = "Current_Funding" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(12)
        Else
            ReturnField = ResponseObject(1)(13)
        End If
    
    ElseIf Field = "Mark_Price" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(15)
        Else
            ReturnField = ResponseObject(1)(16)
        End If

    ElseIf Field = "Open_Interest" Then

        If Timestamp <> "" Then
            ReturnField = ResponseObject(1)(18)
        Else
            ReturnField = ResponseObject(1)(19)
        End If
        
    End If
    

ElseIf Field = "Last_Price_Hist" Or Field = "First_Price_Hist" Then
        ReturnField = ResponseObject(1)(4)

ElseIf Field = "Mid_Price" Then
    ReturnField = (ResponseObject(1)(1) + ResponseObject(2)(1)) / 2

ElseIf Field = "Platform_Status" Then
    
    If ResponseObject(1) = 1 Then
        ReturnField = "Online"
    Else: ReturnField = "Offline"
    End If


ElseIf Field = "Pos_Size_Long" Or Field = "Pos_Size_Short" Or Field = "Funding_Size" Or Field = "Credits_Size" Or Field = "Credits_Size_Sym" Or Field = "Vol.1d" Or Field = "Vol.7d" Or Field = "Vol.30d" Or Field = "Vwap" Then
    ReturnField = ResponseObject(1)(2)
    
Else
    ReturnField = "Error, Unknown Field"
    
End If

End Function

Function CheckErrors(ApiData)

If ApiData = "[]" Then
    CheckErrors = "Error, API returned an empty array. Check provided arguments"
ElseIf ApiData = "[""error"", 11010, ""ratelimit: error""]" Then
    CheckErrors = "Error, RATE LIMIT"
ElseIf ApiData Like "Cannot GET*" Then
    CheckErrors = "Error, not a valid request. Check provided arguments"
Else
    CheckErrors = "Pass"
End If

End Function

Function CheckErrorReturned(Data)

If Data = "Error, API returned an empty array. Check provided arguments" Then
    CheckErrorReturned = "Error"
ElseIf Data = "Error, RATE LIMIT" Then
    CheckErrorReturned = "Error"
ElseIf ApiData = "Error, not a valid request. Check provided arguments" Then
    CheckErrorReturned = "Error"
Else
    CheckErrorReturned = False
End If

End Function


'Function to convert date to timestamp in milliseconds

Public Function ConvertToTimestamp(dt) As String
    ConvertToTimestamp = DateDiff("s", "1/1/1970 00:00:00", dt) * 1000
End Function

'Function to convert timestamp in milliseconds to Date

Public Function ConvertToDate(ts) As Date
    ConvertToDate = DateAdd("s", ts / 1000, "1/1/1970 00:00:00")
End Function


