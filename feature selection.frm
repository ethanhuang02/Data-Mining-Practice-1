VERSION 5.00
Begin VB.Form R76101120_HW1 
   Caption         =   "R76101120_HW1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11655
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Execute 
      Caption         =   "Execute"
      Height          =   615
      Left            =   7560
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   4920
      Left            =   8760
      TabIndex        =   9
      Top             =   1560
      Width           =   2500
   End
   Begin VB.ListBox List3 
      Height          =   4920
      Left            =   6000
      TabIndex        =   8
      Top             =   1560
      Width           =   2500
   End
   Begin VB.ListBox List1 
      Height          =   4920
      Left            =   3240
      TabIndex        =   7
      Top             =   1560
      Width           =   2500
   End
   Begin VB.CommandButton Read 
      Caption         =   "Load Data"
      Height          =   612
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   1212
   End
   Begin VB.TextBox filename 
      Height          =   264
      Left            =   480
      TabIndex        =   2
      Text            =   "soybean-small"
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox strategy 
      Height          =   300
      ItemData        =   "R76101120.frx":0000
      Left            =   4800
      List            =   "R76101120.frx":000A
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox Result 
      Height          =   4920
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   "U(X,Y)"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "H(X,Y)"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "H(X)"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   252
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "Strategy:"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "File name:"
      Height          =   252
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1812
   End
End
Attribute VB_Name = "R76101120_HW1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataList(46, 35) As Single
Dim selected(34) As Boolean
Dim HxArray(35) As Single
Dim HxyArray(35, 35) As Single
Dim UArray(35, 35) As Single

Private Sub Read_Click()
    '讀檔
    Dim i, j As Integer
    Dim path As String
    Dim temp As Variant

    path = App.path & "\" & filename.Text & ".txt"
    i = 0
    Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, temp
            For j = 0 To 35
                Dim tempArr As Variant
                tempArr = Split(temp, ",")
                If j = 35 Then
                    DataList(i, j) = CSng(Split(tempArr(j), "D")(1))
                Else
                    DataList(i, j) = CSng(tempArr(j))
                End If
            Next j
            i = i + 1
        Loop
    Close #1
    
    '計算H(X)
    For i = 0 To 35
        Hx i
        List1.AddItem (i + 1 & " | " & HxArray(i))
    Next i
    
    '計算H(X,Y)
    For i = 0 To 35
        For j = i To 35
            Hxy i, j
            List3.AddItem (i + 1 & "&" & j + 1 & " | " & HxyArray(i, j))
        Next
    Next
    
    '計算U(XY)
    For i = 0 To 35
        For j = i To 35
            U i, j
            List4.AddItem (i + 1 & "&" & j + 1 & " | " & UArray(i, j))
        Next
    Next
    
End Sub
Private Sub Execute_Click()
    Dim index, i, j As Integer
    Dim maxGoodness As Single
    If strategy.Text = "Search Forward" Then

        Result.Clear
        Result.AddItem "Result of Search Forward："
        
        maxGoodness = 0
        '一開始全不選
        For i = 0 To 34
            selected(i) = 0
        Next i

        'forward
        For i = 0 To 34
            index = -1
            For j = 0 To 34
                If Not selected(j) Then
                    selected(j) = 1
                    If Goodness > maxGoodness Then
                        maxGoodness = Goodness
                        index = j
                    End If
                    selected(j) = 0
                End If
            Next
            If index = -1 Then Exit For
            selected(index) = 1

            Result.AddItem ("Attribute Selected：A" & index + 1)
            Result.AddItem ("Goodness：" & maxGoodness)
        Next
        
        Result.AddItem ("-----------------------------")
        
        Result.AddItem ("The Attribute Subset :")
        For i = 0 To 34
            If selected(i) Then
                Result.AddItem ("A" & i + 1)
            End If
        Next i
    End If
    
    If strategy.Text = "Search Backward" Then

        Result.Clear
        Result.AddItem "Result of Search Backward："
        
        '一開始全選
        For i = 0 To 34
            selected(i) = 1
        Next i
        
        maxGoodness = 0
        For i = 0 To 34
            index = -1
            For j = 0 To 34
                If selected(j) Then
                    selected(j) = 0
                    If Goodness > maxGoodness Then
                        maxGoodness = Goodness
                        index = j
                    End If
                    selected(j) = 1
                End If
            Next
            If index = -1 Then Exit For 'goodness下降則停止
            selected(index) = 0

            Result.AddItem ("Attribute Removed：A" & index + 1)
            Result.AddItem ("Goodness：" & maxGoodness)
        Next
        
        Result.AddItem ("-----------------------------")
        Result.AddItem ("The Attribute Subset :")
        For i = 0 To 34
            If selected(i) Then
                Result.AddItem ("A" & i + 1)
            End If
        Next i
    End If
    
End Sub


Static Function Log_2(x) As Single
    If (x = 0) Then
        Log_2 = 0
    Else
        Log_2 = Log(x) / Log(2)
    End If
End Function

Private Function Hx(att)
    Dim i As Integer
    Dim countAtt(6) As Single
    
    For i = 0 To 46
        countAtt(DataList(i, att)) = countAtt(DataList(i, att)) + 1
    Next

    '計算H(x)
    Dim tempHx, p As Single

    For i = 0 To 6
        'List1.AddItem (j & "count" & countarr(j))
        p = countAtt(i) / 47
        tempHx = tempHx + -p * Log_2(p)
        'Hx = Math.Round(Hx + -p * Log_2(p), 3)
    Next
    'List2.AddItem ("H" & col & "｜" & Hx)
    HxArray(att) = tempHx

End Function

Private Function Hxy(att1, att2)
    Dim pxy, tempHxy As Single
    Dim i, j As Integer
    Dim countAtt2(6, 6) As Single
    
    For i = 0 To 46
        countAtt2(DataList(i, att1), DataList(i, att2)) = countAtt2(DataList(i, att1), DataList(i, att2)) + 1
    Next
    
    For i = 0 To 6
        For j = 0 To 6
            pxy = countAtt2(i, j) / 47
            tempHxy = tempHxy + -pxy * Log_2(pxy)
        Next
    Next
    
    HxyArray(att1, att2) = tempHxy

End Function

Private Function U(att1, att2)
    Dim temp As Single
    If (HxArray(att1) + HxArray(att2) = 0) Then
        UArray(att1, att2) = 1
        UArray(att2, att1) = 1
    Else
        temp = 2 * ((HxArray(att1) + HxArray(att2) - HxyArray(att1, att2)) / (HxArray(att1) + HxArray(att2)))
        'temp = Math.Round(2 * ((HxArray(col1) + HxArray(col2) - HxyArray(col1, col2)) / (HxArray(col1) + HxArray(col2))), 6)
        UArray(att1, att2) = temp
        UArray(att2, att1) = temp
    End If
End Function

Private Function Goodness()
    Dim numerator, denominator As Single
    Dim i, j As Integer
    numerator = 0
    denominator = 0
    
    '分子
    For i = 0 To 34
        If selected(i) Then '被選擇
            numerator = numerator + UArray(i, 35)
        End If
    Next
    
    '分母
    For i = 0 To 34
        For j = 0 To 34
            If selected(i) And selected(j) Then
                denominator = denominator + UArray(i, j)
            End If
        Next
    Next
    denominator = Sqr(denominator)
    
    'goodness
    'If denominator <> 0 Then
        Goodness = numerator / denominator
    'End If
    
End Function
