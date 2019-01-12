Attribute VB_Name = "Module1"
Sub TidyUp()
Attribute TidyUp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TidyUp Macro
'

'
    Rows("1:12").Select
    Selection.Delete Shift:=xlUp
    Range("A3").Select
    ActiveCell.FormulaR1C1 = _
        "<a href=""/constituency/aberdeen-north"">Aberdeen North</a>"
    Range("A2").Select
    ActiveSheet.Paste
    Rows("3:24").Select
    Selection.Delete Shift:=xlUp
    Range("A3").Select
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    Range("A5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C1").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C2").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D2").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E2").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F1").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F2").Select
    ActiveSheet.Paste
    Rows("3:22").Select
    Range("A22").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
    Range("A3").Select
    Selection.Copy
    Range("G1").Select
    ActiveSheet.Paste
    Range("A5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G2").Select
    ActiveSheet.Paste
    Range("A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H1").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I1").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I2").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J2").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K1").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K2").Select
    ActiveSheet.Paste
    Rows("3:21").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A4").Select
    Selection.Copy
    Range("L1").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L2").Select
    ActiveSheet.Paste
    Range("A8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M1").Select
    ActiveSheet.Paste
    Range("A10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M2").Select
    ActiveSheet.Paste
    Range("N1").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N2").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O1").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O2").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total"
    Range("A22").Select
    Selection.Copy
    Range("P2").Select
    ActiveSheet.Paste
    Rows("3:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A4").Select
    Selection.Copy
    Range("Q1").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q2").Select
    ActiveSheet.Paste
    Range("A8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R1").Select
    ActiveSheet.Paste
    Range("A10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R2").Select
    ActiveSheet.Paste
    Rows("4:11").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
    Range("A6").Select
    Selection.Copy
    Range("A5").Select
    ActiveSheet.Paste
    Range("A8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B4").Select
    ActiveSheet.Paste
    Range("A10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B5").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C5").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D4").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D5").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E4").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E5").Select
    ActiveSheet.Paste
    Rows("6:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A7").Select
    Selection.Copy
    Range("F4").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F5").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G5").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H5").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I4").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I5").Select
    ActiveSheet.Paste
    Rows("6:21").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A7").Select
    Selection.Copy
    Range("J4").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J5").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K4").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K5").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L5").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M5").Select
    ActiveSheet.Paste
    Rows("7:21").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A8").Select
    Selection.Copy
    Range("N4").Select
    ActiveSheet.Paste
    Range("A10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N5").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O4").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O5").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P4").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P5").Select
    ActiveSheet.Paste
    Range("Q4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "7/5/2019"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "'5-7"
    Range("A22").Select
    Selection.Copy
    Range("Q5").Select
    ActiveSheet.Paste
    Rows("6:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "'8-9"
    Range("S4").Select
    ActiveCell.FormulaR1C1 = "'10-14"
    Range("A9").Select
    Selection.Copy
    Range("R5").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S5").Select
    ActiveSheet.Paste
    Rows("7:13").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("7:7").Select
    Selection.Delete Shift:=xlUp
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Range("A10").Select
    Selection.Copy
    Range("B7").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B8").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C7").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C8").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D7").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D8").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E7").Select
    ActiveSheet.Paste
    Rows("9:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A10").Select
    Selection.Copy
    Range("E8").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F7").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F8").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G7").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G8").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H7").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H8").Select
    ActiveSheet.Paste
    Rows("9:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A10").Select
    Selection.Copy
    Range("I7").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I8").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J7").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J8").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K7").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K8").Select
    ActiveSheet.Paste
    Rows("9:20").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A10").Select
    Selection.Copy
    Range("L7").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L8").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M7").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M8").Select
    ActiveSheet.Paste
    Rows("9:17").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A9").Select
    Selection.Copy
    Range("N7").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N8").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O7").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O8").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P7").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P8").Select
    ActiveSheet.Paste
    Rows("9:19").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A10").Select
    Selection.Copy
    Range("Q7").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q8").Select
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R7").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R8").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S7").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S8").Select
    ActiveSheet.Paste
    Rows("20:20").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("9:19").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Range("A13").Select
    Selection.Copy
    Range("B10").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B11").Select
    ActiveSheet.Paste
    Range("A17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C10").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C11").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D10").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D11").Select
    ActiveSheet.Paste
    Rows("13:23").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A14").Select
    Selection.Copy
    Range("E10").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E11").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F10").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F11").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G10").Select
    ActiveSheet.Paste
    Rows("13:22").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A14").Select
    Selection.Copy
    Range("G11").Select
    ActiveSheet.Paste
    Rows("13:15").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("B13").Select
    Selection.ClearContents
    Rows("14:14").Select
    Selection.Delete Shift:=xlUp
    Range("A16").Select
    ActiveWindow.SmallScroll Down:=12
    Range("A16").Select
    Selection.Copy
    Range("B13").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B14").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C13").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C14").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D13").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D14").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E13").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E14").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F13").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F14").Select
    ActiveSheet.Paste
    Rows("16:34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A17").Select
    Selection.Copy
    Range("G13").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G14").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H13").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H14").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I13").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I14").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J13").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J14").Select
    ActiveSheet.Paste
    Rows("16:32").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A16").Select
    Selection.Copy
    Range("K13").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K14").Select
    ActiveSheet.Paste
    Range("A20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L13").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L14").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M13").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M14").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N13").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N14").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O13").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O14").Select
    ActiveSheet.Paste
    Rows("16:34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A17").Select
    Selection.Copy
    Range("P13").Select
    ActiveSheet.Paste
    Range("A19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P14").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q13").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q14").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R13").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R14").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S13").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S14").Select
    ActiveSheet.Paste
    Rows("15:31").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("17:17").Select
    Selection.Delete Shift:=xlUp
    Range("A19").Select
    Selection.Copy
    Range("B16").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B17").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C16").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C17").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D16").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D17").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E16").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E17").Select
    ActiveSheet.Paste
    Rows("19:33").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A20").Select
    Selection.Copy
    Range("F16").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F17").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G16").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G17").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H16").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H17").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I16").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I17").Select
    ActiveSheet.Paste
    Rows("18:34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A19").Select
    Selection.Copy
    Range("J16").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J17").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K16").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K17").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L16").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L17").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M16").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M17").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=1
    Rows("19:33").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A20").Select
    Selection.Copy
    Range("N16").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N17").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O16").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O17").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P16").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P17").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q16").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q17").Select
    ActiveSheet.Paste
    Rows("19:34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A20").Select
    Selection.Copy
    Range("R16").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R17").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S16").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S17").Select
    ActiveSheet.Paste
    Rows("18:27").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("19:19").Select
    Selection.Delete Shift:=xlUp
    Range("A21").Select
    Selection.Copy
    Range("B18").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B19").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C18").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C19").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D18").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D19").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E18").Select
    ActiveSheet.Paste
    Range("A35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E19").Select
    ActiveSheet.Paste
    Rows("20:35").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A21").Select
    Selection.Copy
    Range("F18").Select
    ActiveSheet.Paste
    Range("A23").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F19").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G18").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G19").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H18").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H19").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I18").Select
    ActiveSheet.Paste
    Range("A35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I19").Select
    ActiveSheet.Paste
    Rows("21:35").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A22").Select
    Selection.Copy
    Range("J18").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J19").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K18").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K19").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L18").Select
    ActiveSheet.Paste
    Rows("21:30").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A22").Select
    Selection.Copy
    Range("L19").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M18").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M19").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N18").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N19").Select
    ActiveSheet.Paste
    Rows("21:30").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A22").Select
    Selection.Copy
    Range("O18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O18").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O19").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P18").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("P19").Select
    ActiveSheet.Paste
    Rows("21:28").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=3
    Range("A22").Select
    Selection.Copy
    Range("Q18").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q19").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R18").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R19").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S18").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S19").Select
    ActiveSheet.Paste
    Rows("21:33").Select
    Range("A33").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("22:22").Select
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
    Range("A24").Select
    Selection.Copy
    Range("B21").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B22").Select
    ActiveSheet.Paste
    Range("A28").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C21").Select
    ActiveSheet.Paste
    Range("A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C22").Select
    ActiveSheet.Paste
    Range("A32").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D21").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D22").Select
    ActiveSheet.Paste
    Rows("24:34").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A25").Select
    Selection.Copy
    Range("E21").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E22").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F21").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F22").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G21").Select
    ActiveSheet.Paste
    Range("A35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G22").Select
    ActiveSheet.Paste
    Rows("24:35").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A25").Select
    Selection.Copy
    Range("H21").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H22").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Rows("24:250").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 1
    ActiveWindow.SmallScroll Down:=2
    Range("A21:H22").Select
    Selection.Cut
    Range("T18").Select
    ActiveSheet.Paste
    Range("A18:AA19").Select
    Selection.Cut
    Range("T16:T17").Select
    ActiveSheet.Paste
    Range("A16:AT17").Select
    Selection.Cut
    Range("T13").Select
    ActiveSheet.Paste
    Range("A13:A14").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H10").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-8
    Rows("10:11").Select
    Selection.Cut
    Range("A10:BT11").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("T7").Select
    ActiveSheet.Paste
    Range("A7:CM8").Select
    Selection.Cut
    Range("T4").Select
    ActiveSheet.Paste
    Range("A4:DF5").Select
    Selection.Cut
    Range("S1").Select
    ActiveSheet.Paste
    ActiveWorkbook.Save
End Sub
