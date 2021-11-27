Attribute VB_Name = "MultiLoopProgressBar"
Option Explicit

Public PROCESS_CANCEL As Boolean
Public ALL_LOOP_BLOCK_CNT As Integer
Dim CUR_LOOP_BLOCK_NO As Integer
Dim CUR_LOOP_BLOCK_BEGIN_PERCENT As Single
Dim CUR_LOOP_BLOCK_END_PERCENT As Single
Dim CUR_LOOP_BLOCK_PERCENT_RANGE As Single

Public Function ProgressBarInitializer(allLoopBlockCnt As Integer)
    
    PROCESS_CANCEL = False
    ALL_LOOP_BLOCK_CNT = allLoopBlockCnt
    CUR_LOOP_BLOCK_NO = 0
    CUR_LOOP_BLOCK_PERCENT_RANGE = 1 / ALL_LOOP_BLOCK_CNT
    ProgressBarForm.Show vbModeless
    ProgressBarForm.percent.BackStyle = fmBackStyleTransparent
    ProgressBarForm.Label1.Width = 0
    
End Function
Public Function ProgressBarManagementFuncBeforeLoop() As Boolean

    If PROCESS_CANCEL = True Then
        Unload ProgressBarForm
        ProgressBarManagementFuncBeforeLoop = False
        Exit Function
    End If
    
    CUR_LOOP_BLOCK_NO = CUR_LOOP_BLOCK_NO + 1
    CUR_LOOP_BLOCK_BEGIN_PERCENT = (CUR_LOOP_BLOCK_NO - 1) / ALL_LOOP_BLOCK_CNT
    CUR_LOOP_BLOCK_END_PERCENT = CUR_LOOP_BLOCK_NO / ALL_LOOP_BLOCK_CNT
    ProgressBarManagementFuncBeforeLoop = True
    
End Function

Public Function ProgressBarManagementFuncInsideLoop(curProgressPercent As Single) As Boolean
    Dim progressPercent As Single
    
    If PROCESS_CANCEL = True Then
        Unload ProgressBarForm
        ProgressBarManagementFuncInsideLoop = False
        Exit Function
    End If
    
    progressPercent = CUR_LOOP_BLOCK_PERCENT_RANGE * curProgressPercent
    ProgressBarForm.Label1.Width = ProgressBarForm.Frame1.Width * progressPercent + (ProgressBarForm.Frame1.Width * CUR_LOOP_BLOCK_BEGIN_PERCENT)  ' ÉoÅ[ÇêiÇﬂÇÈ
    ProgressBarForm.percent.Caption = WorksheetFunction.Round(progressPercent * 100 + CUR_LOOP_BLOCK_BEGIN_PERCENT * 100, 0) & " %"
    DoEvents
    ProgressBarManagementFuncInsideLoop = True
    
End Function

Public Function ProgressBarFinalizer()

    Unload ProgressBarForm
    
End Function
