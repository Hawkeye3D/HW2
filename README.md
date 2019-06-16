# HW2
This is a link to the spreadsheet, should one care to look at it. The images provided show the gist of the work that was accomplished.
[My Week2 Homework](https://drive.google.com/open?id=1iP9Lr1trFYDopMwKk7z_ul6OABPQjwXZ)
## Homework Submission ##
### Additional Features ###
I have included a lookup table that I scoured from Yahoo.  I use it to add clarity to the tickers, what companies they represent and the category of business they are involved with. I have also included a few extra columns, one for variance and one for the approximate trade volume in *dollars* based upon the idea that the average trade per day is the average between the high and the low. On any given day that may not be true, but over the 252 days of trading the **Central-Limit-Theorem** suggest that it is likely.  On the linked spreadsheet there is a pivot table under Charts worksheet for one of the years which is setup to display a chart.  The chart can be flipped between Total Volume and ApproxDollar Volume.  They do not even come close to mirroring one another.  The dominant trading is in Oil related categories; it is an interesting Bar chart.

```vba
 
Option Explicit
Public Applicationinstance As New Excel.Application
 '
 'For Grand Summaries
  Dim MMaxChange As Double, MMinChange As Double, MTotalVolume As Double
  Dim MMaxTick As String, MMinTick As String, MMVoltick As String
  Dim MTotalExchange As Double, MMexchangeTick As String

Sub AnalizeData()
'Declare the variables to be used in the loop
Dim WB As Workbook
Dim Ws As Worksheet
Dim WLookup As Worksheet
Set WB = Application.ActiveWorkbook
Set WLookup = WB.Worksheets("NYQ_Lookup")
Set Ws = Application.ActiveSheet
Application.ScreenUpdating = False

Dim i As Integer
Dim InitTicker As String
Dim lastticker As String
Dim lastname As String
Dim lastcategory As String
Dim StartValue As Double
Dim LastValue As Double
Dim LastClose As Double
Dim MaxValue As Double
Dim MinValue As Double
Dim Variance As Double
Dim percentchange As Double
Dim TotalChange As Double
Dim BestStock As Double
Dim WorstStock As Single
Dim TotalVolume As Double
Dim ApproxTotalExchange As Double
Dim RowPos As Long
Dim TickerCnt As Integer 'count of stocks
Dim HdrArray As Variant
Dim SummaryArray As Variant
'Dim MMaxChange As Double, MMinChange As Double, MTotalValume As Double
'Initialize some values
 MMaxChange = -1000000000
 MMinChange = 100000000
 MTotalVolume = 0
'Create the new Header Array
HdrArray = VBA.Array("<Ticker>", "<name>", "<Category>", "MaxTrade", "MinTrade", "Year Spread", "Percentage Chg", "Total Stock Volume", "Dollar Volume")
Range("j1:p1").Value = HdrArray 'put it on the active worksheet

'Create the New Summary Array and make it vertical
SummaryArray = VBA.Array("Greatest Increase", "Greatest  Decrease", "Greatest Total Volume", "Greatest Dollar Volume")
SummaryArray = Application.WorksheetFunction.Transpose(SummaryArray)
Range("r6:S12").Value = SummaryArray 'Put it on the active worksheet

'
'Init Row Position
'
RowPos = 1
 
lastticker = Cells(RowPos, 3).Value 'open value

'I picked up the lookup data from Yahoo.

'lastname = Application.VLookup(lastticker,  , 2)
'lastcategory = Application.VLookup(lastticker, NYQ_lookup.Range("A1:C3385"), 3)

 
 
'Loop is choosen because it is easier to manage
'finding the end of the dataloop as opposed to using
' a for loop which has to specify a to NUMBER.  VBA does
'have some structures to deal with that, I just don't want to track it down

'Outer loop is a reset loop for dealing with changes in the ticker
RowPos = 2
Do Until Cells(RowPos, 1).Value = ""   ' loop until cell value = 0
 TickerCnt = TickerCnt + 1 ' update ticker count every time logic flows past this point
  lastticker = Cells(RowPos, 1).Value
  StartValue = Cells(RowPos, 6).Value 'this is defined as the first value
  MaxValue = Cells(RowPos, 4).Value 'high value for the day
  MinValue = Cells(RowPos, 5).Value 'low value for the day

'This loop is for iterating through the ticker value and summarizing data for that ticker
Do Until Cells(RowPos, 1).Value <> lastticker Or Cells(RowPos, 1).Value = "" 'monitor ticker until it changes

LastValue = Cells(RowPos, 6).Value 'column 6 is the close column
'The idea behind this is to have a metric of approximately how much money was transacted on the stock over the course of the year
ApproxTotalExchange = ApproxTotalExchange + (Cells(RowPos, 4).Value + Cells(RowPos, 5).Value) / 2 * Cells(RowPos, 7).Value
MaxValue = MaX(Cells(RowPos, 4).Value, MaxValue)
MinValue = Min(Cells(RowPos, 5).Value, MinValue)
TotalVolume = TotalVolume + Cells(RowPos, 7).Value '
RowPos = RowPos + 1 'go to the next row
'If RowPos Mod 10000 = 0 Then MsgBox RowPos
Loop

'Either the ticker has changed value or
'program has come to the end
'so, now we need to save the values and do some calculations
Variance = MaxValue - MinValue
TotalChange = (LastValue - StartValue)
If StartValue <= 0.1 Then
percentchange = 0
Else
percentchange = (LastValue - StartValue) / StartValue * 100
End If

'Now Populate the row
'lastticker,lastname,lastcategory, MaxValue, MinValue,TotalChange,percentchange,TOtalvaluem,ApproxTotalExchange
 UpdateRow TickerCnt + 1, 10, lastticker, lastname, lastcategory, MaxValue, MinValue, TotalChange, percentchange, TotalVolume, ApproxTotalExchange
TotalVolume = 0 'rezero the total value
ApproxTotalExchange = 0 'reset for next summary
Loop
'Lastly. update the Grand Max/Min Stuff ' Dim comments are so I can see what the names are ----
  'Dim MMaxChange As Double, MMinChange As Double, MTotalVolume As Double
  'Dim MMaxTick As String, MMinTick As String, MMVoltick As String
  'Dim MTotalExchange As Double, MMexchangeTick As String
  Cells(6, 20).Value = MMaxTick: Cells(6, 21).Value = MMaxChange
  Cells(7, 20).Value = MMinTick: Cells(7, 21).Value = MMinChange
  Cells(8, 20).Value = MMVoltick: Cells(8, 21).Value = MTotalVolume
  Cells(9, 20).Value = MMexchangeTick: Cells(9, 21).Value = MTotalExchange
  
'allow the screen to be updated
Application.ScreenUpdating = True
'sResult = Application.VLookup(LastTicker, NYQ_Lookup!A1:C3385, 2)

End Sub

Sub UpdateRow(ByVal R As Long, ByVal C As Long, ByVal lastticker As String, ByVal lastname As String _
, ByVal lastcategory As String, ByVal MaxValue As Double, ByVal MinValue As Double, ByVal TotalChange As Double, ByVal percentchange As Double _
, ByVal TotalVolume As Double, ByVal ApproxTotalExchange As Double)
Dim i As Integer
i = 0 'column iterator
Cells(R, C + i).Value = lastticker: i = i + 1
'Cells(R, C + i).Value = lastname: i = i + 1
'Cells(R, C + i).Value = lastcategory: i = i + 1
i = i + 2
Cells(R, C + i).Value = MaxValue: i = i + 1
Cells(R, C + i).Value = MinValue: i = i + 1
Cells(R, C + i).Value = TotalChange: i = i + 1
Cells(R, C + i).Value = percentchange

      If percentchange < 0 Then
       Cells(R, C + i).Interior.ColorIndex = 3
      Else
       Cells(R, C + i).Interior.ColorIndex = 10
      End If
 '      Dim MMaxChange As Double, MMinChange As Double, MTotalVolume As Double
 ' Dim MMaxTick As String, MMinTick As String, MMVol As String
 
 'Get the max and min values
 
 Dim chkit As Double
 chkit = MMaxChange
 MMaxChange = MaX(percentchange, MMaxChange)
 If chkit <> MMaxChange Then MMaxTick = lastticker
 
 chkit = MMinChange
 MMinChange = Min(percentchange, MMinChange)
 If chkit <> MMinChange Then MMinTick = lastticker
      i = i + 1

Cells(R, C + i).Value = TotalVolume
i = i + 1
 chkit = MTotalVolume
 MTotalVolume = MaX(TotalVolume, MTotalVolume)
 If MTotalVolume <> chkit Then MMVoltick = lastticker

Cells(R, C + i).Value = ApproxTotalExchange
 chkit = MTotalExchange
 MTotalExchange = MaX(ApproxTotalExchange, MTotalExchange)
 If MTotalExchange <> chkit Then MMexchangeTick = lastticker
i = i + 1
 


End Sub
'
'Return the highest of two numbers
'
Function MaX(ByVal Currentvalue As Double, Comp As Double) As Double
      If Currentvalue > Comp Then
        MaX = Currentvalue
      Else
        MaX = Comp
      End If
End Function

'Return the Lowest of two numbers
Function Min(ByVal Currentvalue As Double, Comp As Double) As Double
      If Currentvalue < Comp Then
        Min = Currentvalue
      Else
        Min = Comp
      End If
End Function
 
```

