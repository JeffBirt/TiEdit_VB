Option Strict On

Public Class detokenize


#Region "Initialation"

    '' Public Sub New, This subroutine runs when class is loaded
    '  and initializes the delegate list
    Public Sub New()
        MyBase.New()
        dgList(0) = AddressOf funNothingToken
        dgList(1) = AddressOf funNegExpression
        dgList(2) = AddressOf funBaseTokenE3
        dgList(3) = AddressOf funBaseTokenE4
        dgList(4) = AddressOf funBaseTokenE5
        dgList(5) = AddressOf funMultiParam
        dgList(6) = AddressOf funBaseTokenDA
        dgList(7) = AddressOf funBaseTokenE8
        dgList(8) = AddressOf funBaseToken2D
        dgList(9) = AddressOf funBaseToken00
        dgList(10) = AddressOf funBaseToken1C
        dgList(11) = AddressOf funBaseTokenInteger
        dgList(12) = AddressOf funBaseTokenFraction
        dgList(13) = AddressOf funBaseTokenF0
        dgList(14) = AddressOf funBaseTokenD9
        dgList(15) = AddressOf funFunctionToken0F
        dgList(16) = AddressOf funProgramToken3B
        dgList(17) = AddressOf funBaseExprOper
        dgList(18) = AddressOf funBaseTokenFloating
        dgList(19) = AddressOf funProgramToken15
        dgList(20) = AddressOf funProgramToken86
        dgList(21) = AddressOf funVarParameter
        dgList(22) = AddressOf funFunctionToken0A
        dgList(23) = AddressOf funFunctionToken14
        dgList(24) = AddressOf funBaseTokenDC
        dgList(26) = AddressOf funBaseExpr1Expr2
        dgList(27) = AddressOf funBaseExpr2Expr1
    End Sub

#End Region


#Region "Declarations"

    '' OK-Declares gloabal variables
    '  
    Delegate Function dg(ByVal sFun As String) As String    ' new instaace of delgate
    Private dgList(255) As dg                   ' delegate array
    Private dgFun As dg                         ' istance of dg
    Private byteProgramLine() As Byte           ' program line being detokenized
    Private byteRawProgram() As Byte            ' holds raw program
    Private intTokenMode As Integer = 0         ' 0=base, 1=system, 2=function, 3=program
    Private boolMatrixMode As Boolean = False   ' matrix mode flag
    Private boolIsVarString As Boolean = False  ' variable string flag
    Private sLineSuffix As String = ""          ' holds line suffix, i.e "Then"
    Private intLineArrayPointer As Integer = 0  ' aray pointer for program
    Private boolQuoteMode As Boolean = False    ' quote mode flag

    Private sPath As String                     ' holds path of current file
    Private sFolderName As String               ' holds folder name
    Private sVariableName As String             ' set var name
    Private sExtension As String                ' holds extension of current file
    Private sProgramText As String              ' holds text version of program

    Private sCalcName As String                 ' holds calculator name
    Private byteSeperator1() As Byte = {&H8, &H0}       ' holds seperator 1
    Private sProgComment As String              ' holds comment text
    Private byteSeperator2() As Byte = {&H1, &H52, &H0, &H0, &H0} ' holds sep 2, non-tokenized
    Private byteFileType As Byte                ' &H12=prg, &H13=function
    Private byteSeperator3() As Byte = {&H0, &H0, &H0}  ' holds sep 3
    Private intFileLength As Integer            ' holds length of file
    Private byteSeperator4() As Byte = {&HA5, &H5A, &H0, &H0, &H0, &H0} ' holds seperator 4
    Private intProgramSize As Integer           ' holds length of text portion of program
    Private byteSeperator5() As Byte = {&H0, &H1}   ' holds seperator 5
    Private byteProgTag() As Byte = {&H19, &HE4}    ' Program token pair
    Private byteFuncTag() As Byte = {&H17, &HE4}    ' Function token pair
    Const byteEndTag As Byte = &HE5                 ' end tag token
    Private byteSeperator6() As Byte = {&H0, &H0, &H8}    ' holds seperator 6
    Private boolTokenized As Boolean            ' true=tokenized, false=not-tokenized
    Const byteEndMark As Byte = &HDC            ' end mark
    Private intCheckSum As Integer              ' holds check sum

    Private sArgumentList As String             ' hold argumrnt list


#Region "intBaseToken - sBaseToken"
    Private intBaseToken() As Integer = {9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 11, _
                                            11, 12, 12, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8, 0, 21, _
                                            21, 21, 0, 0, 0, 21, 21, 21, 0, 0, 0, 21, 21, 21, 0, 0, _
                                            0, 0, 0, 0, 21, 21, 21, 0, 0, 0, 0, 21, 21, 21, 21, 21, _
                                            21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, _
                                            21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, _
                                            21, 21, 21, 21, 21, 21, 17, 17, 17, 21, 1, 0, 0, 0, 0, 0, _
                                            27, 26, 26, 26, 26, 26, 26, 26, 26, 26, 26, 27, 27, 27, 27, 27, _
                                            27, 27, 27, 26, 26, 21, 21, 21, 21, 21, 21, 21, 21, 0, 21, 21, _
                                            21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, _
                                            21, 0, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 0, 0, 21, 0, _
                                            21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, _
                                            21, 21, 21, 21, 21, 21, 21, 21, 21, 14, 6, 0, 24, 0, 0, 0, _
                                            0, 0, 0, 2, 3, 4, 0, 7, 7, 0, 0, 0, 0, 0, 0, 0, _
                                            13, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}

    Private sBaseToken() As String = {"", "q", "r", "s", "t", "u", "v", "w", _
                                        "x", "y", "z", "a", "b", "c", "d", "e", _
                                        "f", "g", "h", "i", "j", "k", "l", "m", _
                                        "n", "o", "p", "q", "", "@xxx", "@nxxx", "", _
                                        "-", "", "-", "", "Œ", "^", "¢", "-¸", _
                                        "¸", "UND", "undef", "false", "true", "", "UND", "acosh", _
                                        "asinh", "atanh", "UND", "UND", "UND", "cosh", "sinh", "tanh", _
                                        "UND", "UND", "UND", "acos", "asin", "atan", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "cos", "sin", "tan", "UND", _
                                        "UND", "UND", "UND", "abs", "angle", "ceiling", "floor", "int", _
                                        "sign", "sqrt", "^", "ln", "log", "fPart", "iPart", "conj", _
                                        "imag", "real", "approx", "tExpand", "tCollect", "getDenom", "getNum", "UND", _
                                        "cumSum", "det", "colNorm", "rowNorm", "norm", "mean", "median", "product", _
                                        "stdDev", "sum", "variance", "unitV", "Dim", "mat->list", "newList", "rref", _
                                        "ref", "identity", "diag", "colDim", "rowDim", "™", "!", "%", _
                                        Chr(152), "not", Chr(&HAA), "STRUCT VEC: Polar", "STRUCT VEC:vecCylind", "STRUCT VEC:vecSphere", "UND", "Combine", _
                                        Chr(187), "EXPR|CONDITION", " xor ", " or ", " and ", "<", "œ", "=", _
                                        "ž", ">", "", "+", ".+", "-", ".-", "*", _
                                        ".*", "/", "./", "^", ".^", "UND", "solve", "cSolve", _
                                        "nSolve", "zeros", "cZeros", "fMin", "fMax", "UND", "polyEval", "randPoly", _
                                        "crrossP", "dotP", "gcd", "lcm", "mod", "intDiv", "remain", "nCr", _
                                        "nPr", "PÐRx", "PÐRx", "RÐP_THETA_", "RÐPr", "augment", "newMat", "randMat", _
                                        "simult", "UND", "expÐlist", "randNorm", "mRow", "rowAdd", "rowSwap", _
                                        "arcLen", "nInt", Chr(150), "_SIGMA_", "mRowAdd", "ans", "entry", "exact", _
                                        "UND", "comDenom", "expand", "factor", "cFactor", "_integrate_", "_differentiate_", "avgRC", _
                                        "nDeriv", "taylor", "limit", "propFrac", "when", "round", "STRUCT DMSNUMBER", "left", _
                                        "right", "mid", "shift", "seq", "list", "subMat", "matrix", "rand", _
                                        "min", "max", "", "", "TYPE/STRUCT matrix", "TYPE/STRUCT prog/func", "TYPE/STRUCT data", "TYPE/STRUCT GDB", _
                                        "TYPE/STRUCT picture", "TYPE/STRUCT text", "TYPE/STRUCT figure", "TYPE/STRUCT macro", "", "", "", Chr(&HA8), _
                                        ":", vbCrLf, "", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND"}
#End Region


#Region "intSystemToken - sSystemToken"

    Private sSystemToken() As String = {"UND", "x_bar_", "y_bar", "_SIGMA_x", "_SIGMA_x_^2", "_SIGMA_y", "_SIGMA_x_^2", "_SIGMA_xy", "Sx", _
                                        "Sy", "_sigma_x", "sigma_y", "nStat", "minX", "minY", "q1", "medStat", _
                                        "q3", "maxX", "maxY", "corr", "R_^2_", "medx1", "medx2", "medx3", _
                                        "medy1", "medy2", "medy3", "xc", "yc", "zc", "tc", "rc", _
                                        "_THETA_c", "nc", "xfact", "yfact", "zfact", "xmin", "xmax", "xscl", "ymin", "ymax", _
                                        "yscl", "_DELTA_x", "_DELTA_y", "xres", "xgrid", "ygrid", "zmin", "zmax", "zscl", _
                                        "eye_THETA_", "eye_PHI_", "_THETA_min", "_THETA_max", "_THETA_step", "tmin", "tmax", "tstep", _
                                        "nmin", "nmax", "plotStrt", "plotStep", "zxmin", "zxmax", "zxscl", "zymin", _
                                        "zymax", "zyscl", "zxres", "z_THETA_min", "z_THETA_max", "z_THETA_step", "ztmin", "ztmax", "ztstep", _
                                        "zxgrid", "zygrid", "zzmin", "zzmax", "zzscl", "zeye_THETA_", "zeye_PHI_", "znmin", _
                                        "znmax", "zpltstep", "zpltstrt", "seed1", "seed2", "ok", "errornum", "sysMath", _
                                        "sysData", "UND", "regCoef", "tblInput", "tblStart", "_DELTA_tbl", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND"}

    Private intSystemToken() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}

#End Region


#Region "intFunctionToken - sFunctionToken"
    Private sFunctionToken() As String = {"UND", "# STRING-EXPR", "getkey", "getFold", "switch", "UND", "ord", "expr", _
                                        "char", "string", "getType", "getMode", "setFold", "ptTest", "pxlTest", "setGraph ", _
                                        "setTable", "setMode", "format", "inString", "&", "DMSNUMBER |>DD", "EXPR |>Rect", "EXPR |>Polar", _
                                        "EXPR |>Cylind", "EXPR |>Sphere", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "part?", "_THETA_c", "nc", "xfact", "yfact", "zfact", "xmin", "ymin", _
                                        "ymax", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND"}

    Private intFunctionToken() As Integer = {0, 21, 0, 0, 21, 0, 22, 22, 22, 22, 22, 22, 22, 15, 15, 15, _
                                            15, 15, 15, 21, 23, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
#End Region


#Region "intProgramToken"
    Private sProgramToken() As String = {"UND", "ClrDraw", "ClrGraph", "ClrHome", "ClrIO", "ClrTable", "Custom", "Cycle", _
                                        "Dialog", "DispG", "DispTbl", "Else", "EndCustom", "EndDlog", "EndFor", "EndFunc", _
                                        "EndIf", "EndLoop", "EndPrgm", "EndTBar", "EndTry", "EndWhile", "Exit", "Func", _
                                        "Loop", "Prgm", "ShowStat", "Stop", "Then", "Toolbar", "Trace", "Try", _
                                        "ZoomBox", "ZoomData", "ZoomDec", "ZoomFit", "ZoomIn", "ZoomInt", "ZoomOut", "ZoomPrev", _
                                        "ZoomRcl", "ZoomSqr", "ZoomStd", "ZoomSto", "ZooTrig", "DrawFunc", "DrawInv", "Goto ", _
                                        "Lbl ", "Get VAR", "Send LIST", "GetCalc ", "SendCalc ", "NewFold", "PrintObj", "RclGDB", _
                                        "StoGDB", "ElseIf ", "If ", "If ", "RandSeed", "While ", "LineTan", "CopyVar", _
                                        "rename", "Style", "Fill", "Request", "PopUp", "PtChg", "PtOff", "PtOn", _
                                        "PxlChg", "PxlOff", "PxlOn", "MoveVar", "DropDown ", "Output", "PtTtext ", "PxlText ", _
                                        "DrawSlp", "Pause", "Return ", "Input", "PlotsOff", "PlotsOn", "Title ", "Item", _
                                        "InputStr", "Linehorz", "LineVert", "Pxlhorz", "PxlVert", "AndPic", "RclPic", "RplcPic", _
                                        "XorPic", "DrawPol", "Text", "OneVar", "StoPic", "Graph", "Table", "NewPic", _
                                        "DrawParm", "CyclePic", "CubicReg", "ExpReg", "LinReg", "LnReg", "MedMed", "PowerReg", _
                                        "QuadReg", "QuartReg", "TwoVar", "Shade", "For ", "Circle", "PxlCrcl", "NewPlot", _
                                        "Line", "PxlLine", "Disp ", "FnOff ", "FnOn ", "Local ", "DeFold", "DelVar ", _
                                        "Lock", "Prompt", "SortA", "SortD", "UnLock", "NewData", "Define ", "Else", _
                                        "ClrErr", "PassErr", "DispHome", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND", _
                                        "UND", "UND", "UND", "UND", "UND", "UND", "UND", "UND"}

    Private intProgramToken() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 19, 0, _
                                            0, 19, 0, 0, 19, 19, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 16, 0, 16, 0, 0, 0, 0, _
                                            0, 0, 0, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 5, 0, 0, 0, 0, 0, 0, 0, 0, 5, 21, 0, _
                                            0, 0, 0, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}

#End Region


#End Region


#Region "Properties"

    '' OK-This Property sets byteProgramLine
    '  returns nothing, used for frmDetokenize
    Public WriteOnly Property setProgLine() As Byte()
        Set(ByVal Value As Byte())
            byteProgramLine = Value
            intLineArrayPointer = 0
        End Set
    End Property

    '' This Property gets or sets current program file, setting causes raw program
    '  to be split into it's indivual data parts, and the program to be detokenized
    Public Property calcProgFile() As Byte()
        Get
            Dim sParsedProgram As String = funParseText(sProgramText)   ' parse raw text file

            Return byteRawProgram               ' returns current raw program saved by class
        End Get

        Set(ByVal value() As Byte)              ' used to set now raw program
            ReDim byteRawProgram(value.Length)  ' resize array to hold program
            byteRawProgram = value              ' save raw program
            funSplit()                          ' split raw program into discrete data
        End Set
    End Property

    '' This Property sets or gets sProgramText, sProgramText holds
    '  text verison of program
    Public Property ProgramText() As String
        Get
            Return sProgramText
        End Get

        Set(ByVal sText As String)
            sProgramText = sText
        End Set
    End Property

    '' This property gets or sets name of calculator that wrote the program
    '  returns current value if existing or sets based on file type if new file
    Public Property calcName() As String
        Get
            Return sCalcName        ' send back calculator name
        End Get

        Set(ByVal sName As String)
            Dim intNameLength As Integer = sName.Length ' get length of string sent
            If intNameLength < 8 Then                   ' if less than eight charecters
                sName.PadRight(8, CChar("*"))           ' pad with astericks
            ElseIf intNameLength > 8 Then               ' if its longer than eight charecters
                sName = sName.Substring(0, 7)           ' than only use first eight charecters
            End If                                      ' 
            sCalcName = sName                           ' transfer calc name to sCalcName
        End Set
    End Property



    '' This property gets or sets path
    '  when path is set, the full path\folder.variable.ext is split into it's parts
    Public Property path() As String

        Get
            Return sPath
        End Get

        Set(ByVal sValue As String)

            Dim charFileHash As Char() = {"\"c}                 ' convert "\" to char array
            Dim charDot As Char() = {"."c}                      ' convert "." to char array
            Dim saPath As String()                              ' holds whole file name
            Dim intLength As Integer                            ' holds # of split segments

            saPath = sValue.Split(charFileHash)                 ' split path at hash marks
            intLength = saPath.GetLength(0) - 1                 ' number of resulting segemnts -1
            '                                                   ' which points to file name
            Dim sFileName As String = saPath(intLength)         ' gets file name (folder.variable.ext)
            sPath = sValue.TrimEnd(sFileName.ToCharArray)       ' strip file name from path save to sPath
            saPath = sFileName.Split(charDot)                   ' split file name at dots
            intLength = saPath.GetLength(0) - 1                 ' number of resulting segments -1
            sFolderName = saPath(0)
            sVariableName = saPath(1)
            sExtension = saPath(2)                      ' which points to extension

        End Set

    End Property

    '' This property gets or sets the Folder Name
    '  returns folder name in string, sets folder to main
    Public Property folderName() As String
        Get                                 '
            Return sFolderName              ' get folder name
        End Get                             '

        Set(ByVal sFolder As String)                        '
            Dim intFolderLength As Integer = sFolder.Length ' get length of string sent
            If intFolderLength > 8 Then                     ' if its longer than eight charecters
                sFolder = sFolder.Substring(0, 7)           ' than only use first eight charecters
            End If                                          '
            sFolderName = sFolder           ' set folder name
        End Set                             '
    End Property

    '' This Property gets or sets Variable Name
    '  returns string, set value must be 8 charecters or less
    Public Property variableName() As String
        Get                                             '
            Return sVariableName                        ' return variable name
        End Get                                         '

        Set(ByVal sVariable As String)                  '
            Dim intVariableLength As Integer = sVariable.Length ' get length of string sent
            If intVariableLength > 8 Then               ' if its longer than eight charecters
                sVariable = sVariable.Substring(0, 7)   ' than only use first eight charecters
            End If                                      '
            sVariableName = sVariable                   ' and transfer name to byteRawProgram
        End Set                                         '

    End Property

    '' This property gets extension of current file
    '
    Public Property extension() As String
        Get
            Return sExtension
        End Get
        Set(ByVal sValue As String)
            sExtension = sValue
        End Set
    End Property



    '' This property gets or sets Program file text
    '  returns string maximum of 40 charecters
    Public Property progComment() As String
        Get
            Return sProgComment
        End Get

        Set(ByVal sComment As String)
            Dim intCommentLength As Integer = sComment.Length   ' get length of string sent
            If intCommentLength > 40 Then                       ' if its longer than 40 charecters
                sComment = sComment.Substring(0, 39)            ' than only use first 40 charecters
            End If
            sProgComment = sComment
        End Set

    End Property

    '' OK-This property gets or sets Seperator 1
    '  returns and accepts a byteArray of size 2
    Public Property Seperator1() As Byte()
        Get                                 '
            Return byteSeperator1           ' return current seperator 1
        End Get                             '

        Set(ByVal byteSep1() As Byte)       '
            byteSeperator1 = byteSep1       ' set seperator one with new value
        End Set                             '
    End Property

    '' OK-This Property gets or sets Seperator 2
    '  returns string, sets 00 01 00 52 00 00 00
    Public Property Seperator2() As Byte()
        Get                                 '
            Return byteSeperator2           ' return current seperator 2
        End Get                             '

        Set(ByVal byteSep2() As Byte)       '
            byteSeperator2 = byteSep2       ' set seperator 2 with new value
        End Set
    End Property

    '' This Property gets or sets Seperator 3
    '  returns string, sets 
    Public Property Seperator3() As Byte()
        Get                                 '
            Return byteSeperator3           ' return current seperator 3
        End Get                             '

        Set(ByVal byteSep3() As Byte)       '
            byteSeperator3 = byteSep3       ' set seperator 3 with new value
        End Set

    End Property

    '' This Property gets or sets Seperator4
    '  returns string
    Public Property Seperator4() As Byte()
        Get                                 '
            Return byteSeperator4           ' return current seperator 4
        End Get                             '

        Set(ByVal byteSep4() As Byte)       '
            byteSeperator4 = byteSep4       ' set seperator 4 with new value
        End Set

    End Property

    '' OK-This Property gets File Length
    '  returns string
    Public ReadOnly Property FileLength() As Integer
        Get
            Return intFileLength
        End Get
    End Property

    '' OK-This Property gets or sets Bytes2Read
    '  returns string
    Public ReadOnly Property Bytes2Read() As String
        Get
            Dim intI As Integer         ' array pointer
            Dim sReturn As String       ' return string
            For intI = 86 To 87         ' Bytes to read
                sReturn += Hex(byteRawProgram(intI)) + " "
            Next
            Return sReturn
        End Get
    End Property

    '' OK-This Property gets or sets CheckSum
    '  returns string
    Public ReadOnly Property CheckSum() As String
        Get
            Dim intI As Integer         ' array pointer
            Dim sReturn As String       ' return string
            Dim intT As Integer = 0     ' holds calculated check sum
            Dim intFileLength As Integer = byteRawProgram.Length            ' location based on file end
            For intI = (intFileLength - 1) To (intFileLength - 3) Step -1   ' Check Sum
                sReturn += Hex(byteRawProgram(intI)) + " "
            Next

            For intI = 86 To byteRawProgram.Length - 4
                intT += byteRawProgram(intI)
            Next

            Return sReturn + Hex(intT)

        End Get
    End Property

    '' OK-This Property gets sets End mark
    '  returns string
    Public ReadOnly Property EndMark() As String
        Get
            Dim intI As Integer         ' array pointer
            Dim sReturn As String       ' return string
            Dim intFileLength As Integer = byteRawProgram.Length    ' location based on file end
            sReturn += Hex(byteRawProgram(intFileLength - 4))       ' End Mark
            Return sReturn
        End Get
    End Property

    '' OK-This Property gets file type program of function
    '  Returns false=function, true=program
    Public Property fileType() As Byte
        Get
            Return byteFileType
        End Get
        Set(ByVal byteType As Byte)
            byteFileType = byteType
        End Set

    End Property

    '' OK-This Property gets tokenization status of file
    '  returns false = tokenized, true=not tokenized
    Public ReadOnly Property notTokenized() As Boolean
        Get
            Dim boolReturn As Boolean       ' return flag
            If byteRawProgram(byteRawProgram.Length - 5) = &H8 Then boolReturn = True
            Return boolReturn               ' if &h08 is found program is not tokenized
        End Get
    End Property

    '' OK-This Property gets Argument List from loaded file
    '  returns string
    Public ReadOnly Property ArgList() As String
        Get
            Dim sReturn As String                       ' return string
            Dim intI As Integer = byteRawProgram.Length - 5
            Dim sComma As String = ""                   ' used to build tokenized arg. list
            Dim intT As Integer = 0                     ' array pointer for arg. list build
            intLineArrayPointer = 0                     ' detokenize routing array pointer
            ReDim byteProgramLine(Me.ArgLength)         ' resize detokinze array

            If Me.notTokenized = False Then             ' if file is tokenized do the following
                While byteRawProgram(intI) <> &HE5      ' until end tag
                    If byteRawProgram(intI) <> &H0 Then ' ignore nulls
                        byteProgramLine(intT) = byteRawProgram(intI)    ' build arg. list array
                        intT += 1                       ' inc arg. list array pointer
                    End If                              '   
                    intI -= 1                           ' dec. program file array pointer
                End While                               '

                While intLineArrayPointer < intT        ' run till end of arg. list
                    sReturn += sComma + funDetokenize() ' detokeize arg. list array
                    sComma = ","                        ' enable comma between args.
                End While                               '

                sReturn = "(" + sReturn + ")"           ' wrap arg. list in parenthsis
            Else                                        ' program is not tokenized so do this
                intT = 88                               ' set progrm starting point
                Do                                      '
                    sReturn += Chr(byteRawProgram(intT)) ' build arg. list string
                    intT += 1                           ' inc array pointer
                Loop Until byteRawProgram(intT) = &H29  ' keep going until closing parenthsis found
                sReturn += ")"                          ' add closing parenthsis
            End If                                      '

            Return sReturn                              ' return arg. list
        End Get

    End Property

    '' OK-This Property gets Argument List Length from loaded file
    '  returns integer
    Public ReadOnly Property ArgLength() As Integer
        Get
            Dim intReturn As Integer                ' return integer
            Dim intI As Integer = byteRawProgram.Length - 5 ' starting point for arg. list serach

            If Me.notTokenized = False Then         ' if program is tokenized then do this
                While byteRawProgram(intI) <> &HE5  ' keep going until end tag
                    If byteRawProgram(intI) <> &H0 Then intReturn += 1 ' ignore nulls
                    intI -= 1                       ' dec array pointer
                End While                           '
            Else                                    ' program is not tokenized so do this
                intReturn = 88                      ' set starting point
                Do                                  ' 
                    intReturn += 1                  ' inc counter 
                Loop Until byteRawProgram(intReturn) = &H29 ' do untill closing parenthsis found
                intReturn -= 87                     ' adjust total for starting point offset
            End If                                  '

            Return intReturn                        ' return arg length
        End Get
    End Property

    '' OK-This Property gets HexDump
    '  returns string
    Public ReadOnly Property HexDump() As String
        Get
            Dim sReturn As String               ' return string
            Dim sTemp As String                 ' temp string
            Dim intI As Integer = byteRawProgram.Length - 1 ' progra length -1, ignore trailing 0

            If Me.notTokenized = False Then     ' if program is tokenized
                While intI >= 88                ' do untill program end
                    sTemp = Hex(byteRawProgram(intI)) + " "        ' get current byte from array
                    If sTemp.Length < 3 Then sTemp = "0" + sTemp '  ' add leading zero if needed
                    sReturn += sTemp                                ' build return sting
                    If byteRawProgram(intI) = &HE8 Then             ' if token is &HE8 "end of line"
                        sReturn += vbCrLf       ' add CrLf to start new line
                    End If                      '
                    intI -= 1
                End While                       '
            Else                                ' program is not tokenized so do this
                Dim intT As Integer = 88        ' starting point
                While intT <= intI              ' go until program end
                    sTemp = Hex(byteRawProgram(intT)) + " "        ' get current byte from array
                    If sTemp.Length < 3 Then sTemp = "0" + sTemp '  ' add leading zero if needed
                    sReturn += sTemp                                ' build return sting
                    If byteRawProgram(intT) = &HD Then sReturn += vbCrLf ' if end of line add CrLf
                    intT += 1                   ' inc array pointer
                End While                       '
            End If                              '

            Return sReturn                      ' return hex dump string
        End Get
    End Property

    '' OK-This Property gets RawDump
    '  returns string, raw unformatted hex dump
    Public ReadOnly Property RawDump() As String

        Get
            Dim sReturn As String                               ' return string
            Dim intI As Integer                                 ' program length
            Dim sTemp As String                                 ' current byteRawProgram(intI)

            For intI = 0 To byteRawProgram.Length - 1           ' do untill program end
                sTemp = Hex(byteRawProgram(intI)) + " "         ' get current byte in array
                If sTemp.Length < 3 Then sTemp = "0" + sTemp '  ' add leading zero if needed
                sReturn += sTemp                                ' build return string
            Next                                                ' loop
            Return sReturn                                      ' return raw hex dump string
        End Get

    End Property


#End Region


#Region "Methods"


#Region "Detokenize"

    '' Function funSplit
    '  takes raw program and splits it into it's discrete data types
    Private Function funSplit() As Boolean

        Dim intI As Integer         ' array pointer
        Dim intT As Integer = 0
        sCalcName = ""              ' clear string
        sFolderName = ""            ' clear string
        sProgComment = ""           ' clear string
        sVariableName = ""          ' clear string
        intFileLength = 0           ' set to known state
        intProgramSize = 0          ' set to known state
        sProgramText = ""           ' clear string
        boolTokenized = True        ' set to known state

        '' get text portions of file
        '
        For intI = 0 To 7           ' get calculator name
            sCalcName += ChrW(byteRawProgram(intI))   ' builds calculator name string
        Next                        '
        For intI = 10 To 16         ' get folder name
            If byteRawProgram(intI) > 0 Then sFolderName += ChrW(byteRawProgram(intI))
        Next                        '
        For intI = 18 To 57         ' Program File Text
            If byteRawProgram(intI) > 0 Then sProgComment += ChrW(byteRawProgram(intI))
        Next
        For intI = 63 To 71         ' get variable name
            If byteRawProgram(intI) > 0 Then sVariableName += ChrW(byteRawProgram(intI))
        Next                        '

        '' get file type, file length, program size, tokenized status
        '
        byteFileType = byteRawProgram(72)           ' get file type

        Dim dblByteMult As Double = 1               ' multiplier
        For intI = 76 To 79                         ' File Length
            intFileLength = CInt((byteRawProgram(intI) * dblByteMult) + intFileLength)
            dblByteMult = dblByteMult * 256         ' * next byte by multiplier add to current value
        Next

        intProgramSize = byteRawProgram(86) * 256 + byteRawProgram(87)  ' program size
        If byteRawProgram(byteRawProgram.Length - 4) = &H8 Then boolTokenized = False

        '' If program is tokenized pull the program body out and save it in byteProgramLine
        '  if it's not tokenized parse ascii and save in sProgram Text
        If boolTokenized = True Then            ' if program is tokenized
            ReDim byteProgramLine(byteRawProgram.Length - 91)   ' resize array to fit size of program
            intI = byteRawProgram.Length - 3    ' start of tokenized program
            intT = 0                            ' array pointer
            Dim bytetemp As Byte
            While intI >= 88                    ' go until end of program
                bytetemp = byteRawProgram(intI)
                byteProgramLine(intT) = byteRawProgram(intI)    ' build array of program tokens
                intI -= 1                       ' dec byteRawProgram array pointer
                intT += 1                       ' dec program array pointer
            End While
            sProgramText = funGetProgramText()  ' detokenize program
        Else                                                ' program is not tokenized so do this
            intT = 88                                       ' start of program
            intI = byteRawProgram.Length - 7                ' end of program
            While intT <= intI                              ' do until end of program
                If byteRawProgram(intT) = &HD Then          ' if end of line tag
                    sProgramText += vbCrLf                  ' add crlf at end of line
                ElseIf byteRawProgram(intT) = &H16 Then     ' if store tag
                    sProgramText += Chr(&HBB)               ' add store charecter
                ElseIf byteRawProgram(intT) = &HAD Then     ' if negitive tag
                    sProgramText += Chr(&HAA)               ' add neg charecter
                Else : sProgramText += Chr(byteRawProgram(intT)) ' otherwise add charecter
                End If                                      ' 
                intT += 1                                   ' inc array pointer
            End While                                       '
            sProgramText = sVariableName + sProgramText     ' add variable name to arg list
        End If                                              '

        Return True                                         ' could add return true=no errors

    End Function

    Public Function testDetokenize() As String

        Dim intProgLength As Integer = byteProgramLine.Length
        Dim sReturn As String
        While intLineArrayPointer < intProgLength       ' run till end of prog. array
            sReturn += funDetokenize()                  ' build return sting with detokenized prg.
        End While
        Return sReturn

    End Function

    '' OK-Public Function Detokenize, detokenizes program held in ByteRawProgram
    '  returns detokenized as string
    Private Function funGetProgramText() As String

        intLineArrayPointer = 0                         ' reset arry pointer
        Dim sReturn As String                           ' return string
        Dim intProgLength As Integer = byteProgramLine.Length   '
        '                                               '
        While intLineArrayPointer < intProgLength       ' run till end of prog. array
            sReturn += funDetokenize()                  ' build return sting with detokenized prg.
        End While                                       '
        '                                               '
        Return sReturn                                  ' return detokenized program as string

    End Function

    '' OK-Function funDetokenize, called by btn1.Click and other functions to
    '  detokenize next token
    Public Function funDetokenize() As String

        Dim sReturn As String                               ' return string
        Dim intToken As Integer                             ' holds token at inLineArrayPointer
        Dim intPointer As Integer                           ' holds delgate pointer value

        intToken = byteProgramLine(intLineArrayPointer)     ' get token at current array location
        intLineArrayPointer += 1                            ' move to next token

        If intTokenMode = 0 Then                            ' if base token
            intPointer = intBaseToken(intToken)             ' gets function pointer
            dgFun = dgList(intPointer)                      ' sets delgate to pointer
            sReturn += dgFun(sBaseToken(intToken))          ' calls delegate
        ElseIf intTokenMode = 1 Then                        ' if system token
            sReturn += sSystemToken(intToken)               ' build token string
            intTokenMode = 0                                ' reset token mode to 0
        ElseIf intTokenMode = 2 Then                        ' if function token
            intTokenMode = 0                                ' reset token mode to 0
            intPointer = intFunctionToken(intToken)         ' gets function pointer
            dgFun = dgList(intPointer)                      ' sets delegate to pointer
            sReturn += dgFun(sFunctionToken(intToken))      ' calls delegate
        ElseIf intTokenMode = 3 Then                        ' if program token
            intTokenMode = 0                                ' reset token mode to 0
            intPointer = intProgramToken(intToken)          ' gets function pointer
            dgFun = dgList(intPointer)                      ' sets delgate to pointer
            sReturn += dgFun(sProgramToken(intToken))       ' calls delegate
        End If                                              '

        Return sReturn                                      ' returns detokenized string

    End Function


#End Region


#Region "GeneralToken functions"

    '' OK-Function funNothingToken called for tokens requiring no action
    '  returns token string
    Private Function funNothingToken(ByVal sFun As String) As String
        Return sFun                 ' returns only token string
    End Function

    '' OK-Function funGetStringLength gets length of string at current position in array
    '  returns integer
    Private Function funGetStringLength() As Integer

        Dim intStringLength As Integer = 0  ' return value
        While byteProgramLine(intLineArrayPointer + intStringLength) <> 0   'go till '0' terminator
            intStringLength += 1            ' inc array pointer
        End While                           '
        Return intStringLength              ' return length

    End Function

    '' OK-Function funGetString returns string from current line starting at intStringEnd
    '  given by intStringLength
    Private Function funGetString(ByVal intStringEnd As Integer, ByVal intStringLength As Integer) As String

        Dim sReturn As String       ' return string
        Dim intTemp As Integer      ' holds integer representing current array location
        While intStringLength > 0   ' while not at end of string
            intTemp = CInt(byteProgramLine(intStringLength + intStringEnd - 1)) ' array to integer
            sReturn += Chr(intTemp) ' adds charecter to sReturn
            intStringLength -= 1    ' dec pointer to next charecter
        End While                   '
        Return sReturn              ' return string

    End Function

    '' OK-Function funGetExpression , returns expression staring at current intLineArrayPointer
    '  ignores modifiers such as FO, returns only expressions in sReturn
    Private Function funGetExpression() As String

        Dim SReturn As String = funDetokenize()         ' detokenizes next expression
        If SReturn = "" Then SReturn = funDetokenize() '' if expression is empty try again
        Return SReturn                                  ' return expression as string

    End Function

    '' OK-Function funMultiParam
    '  returns variable parameter functions as string
    Private Function funMultiParam(ByVal sFun As String) As String

        Dim sReturn As String = ""          ' return string
        Dim sComma As String = ""           ' used to add comma between arguments

        While byteProgramLine(intLineArrayPointer) <> &HE5 _
        And byteProgramLine(intLineArrayPointer) <> &HE8        ' if not end tag and not EOL then
            sReturn += sComma + funGetExpression()              ' add net expresion to argument list
            sComma = ","                                        ' enable comma between arguments
        End While                                               '
        Return sFun + sReturn                                          ' enclose parenthsis

    End Function

    '' OK-Function funVarParameter 
    '  returns variable parameter functions as string
    Private Function funVarParameter(ByVal sFun As String) As String

        Dim sReturn As String = ""          ' return string
        Dim sComma As String = ""           ' used to add comma between arguments

        While byteProgramLine(intLineArrayPointer) <> &HE5 _
        And byteProgramLine(intLineArrayPointer) <> &HE8        ' if not end tag and not EOL then
            sReturn += sComma + funGetExpression()              ' add net expresion to argument list
            sComma = ","                                        ' enable comma between arguments
        End While                                               '
        Return sFun + "(" + sReturn + ")"                       ' enclose parenthsis

    End Function

#End Region


#Region "BaseToken Functions"

    '' OK-Returns string ending at current intLineArrayPointer, adds brackets, parenthsis, and quotes
    '  as needed
    Private Function funBaseToken00(ByVal sFun As String) As String

        Dim sReturn As String = ""                                      ' return string
        Dim intStringLength As Integer = funGetStringLength()           ' gets length of string

        sReturn += funGetString(intLineArrayPointer, intStringLength)   ' add string to sReturn
        intLineArrayPointer += intStringLength + 1                      ' bypass &h00 string terminator
        If boolQuoteMode = True Then sReturn = Chr(&H22) + sReturn + Chr(&H22) ' add quotes if needed
        boolQuoteMode = False                                           ' clear quote flag
        Return sReturn                                                  ' return string

    End Function

    '' OK-Function funBaseToken1C, '*system token' sets token mode to 1
    '  so next token will be interpreted as a system token, returns ""
    Private Function funBaseToken1C(ByVal sFun As String) As String
        intTokenMode = 1    ' sets token mode to 1
        Return ""           ' returns nothing
    End Function

    '' OK-Function funBaseTokenInteger, 'STRUCT +- integer', returns +- integer currently pointed to
    '  by intLineArrayPointer, returns integer as string
    Private Function funBaseTokenInteger(ByVal sFun As String) As String

        Dim intTIBinInt As Integer = 0              ' integer taken from array
        Dim intByteMult As Integer = 1              ' multiplier
        Dim intEndofBytes As Integer = byteProgramLine(intLineArrayPointer) + intLineArrayPointer
        intLineArrayPointer += 1                    ' get length of integer inc array pointer

        While intLineArrayPointer <= intEndofBytes  ' go till end of integer bytes 
            intTIBinInt = CInt((byteProgramLine(intLineArrayPointer) * intByteMult) + intTIBinInt)
            intByteMult = intByteMult * 256         ' * next byte by multiplier add to current value
            intLineArrayPointer += 1                ' inc multiplier and inc array pointer
        End While                                   '

        Return sFun + CStr(intTIBinInt)             ' return token string + integer

    End Function

    '' OK-Function funBaseTokenFraction, 'STRUCT +- fraction', returns +- fraction currently pointed to
    '  by intLineArrayPointer, returns fraction as string
    Private Function funBaseTokenFraction(ByVal sFun As String) As String

        Return sFun + funBaseTokenInteger("") + "/" + funBaseTokenInteger("")

    End Function

    '' OK-Function funBaseTokenFloating, 'STRUCT +- floating', returns +- floating currently pointed to
    '  by intLineArrayPointer, returns floating as string
    Private Function funBaseTokenFloating(ByVal sFun As String) As String

        Dim sReturn As String = ""              ' return string
        Dim sSign As String = ""                ' sign string
        Dim intI As Integer                     ' array pointer
        Dim dblFloat As Double                  ' used for multiplier
        Dim intExponent As Integer = 0          ' used to hold exponent
        Dim byteBCD As Byte                     ' holds BCD digit

        For intI = 6 To 0 Step -1               ' do until end of floating point number
            byteBCD = byteProgramLine(intLineArrayPointer + intI)   ' get BCD pair and calc their val.
            sReturn += Chr(CByte(((byteBCD And &HF0) \ 16) + 48)) + Chr(CByte((byteBCD And &HF) + 48))
        Next                                                        ' add to sReturn
        intLineArrayPointer += 7                                    ' bypass floating point number

        intExponent = (byteProgramLine(intLineArrayPointer) And &HFE)   ' get low byte of exponent
        intExponent += byteProgramLine(intLineArrayPointer + 1) * 256   ' get hi byte of exponent
        intExponent = intExponent - 16384 - 13                          ' calculate exponent-adjust for offset
        dblFloat = 10 ^ intExponent                                     ' calcualte actual power of 10
        intLineArrayPointer += 2                                        ' inc past exponent bytes

        dblFloat = (CDbl(sReturn) * dblFloat)                           ' convert string to double
        Return sFun + CStr(dblFloat)                                    ' return token string+Float String

    End Function

    '' OK-Function funBaseToken2D, siganls that the following string should be
    '  enclosed in quotes, returns ""
    Private Function funBaseToken2D(ByVal sFun As String) As String
        boolQuoteMode = True    ' set quote flag to true 
        Return ""               ' does nothing returns nothing
    End Function

    '' OK-Function funBaseExprOper, EXPR Operator
    '  returns
    Private Function funBaseExprOper(ByVal sFun As String) As String

        Return funGetExpression() + sFun    ' return expression then operator i.e. '15%'

    End Function

    '' OK-Function funNegExpression
    '  returns expression as string
    Private Function funNegExpression(ByVal sFun As String) As String
        Return sFun + funGetExpression()
    End Function

    '' OK-Function funBaseExpr1Expr2
    '  
    Private Function funBaseExpr1Expr2(ByVal sFun As String) As String

        Return funGetExpression() + sFun + funGetExpression()   ' return expr1+op+expr2

    End Function

    '' OK-Function funBaseExpr2Expr1
    '
    Private Function funBaseExpr2Expr1(ByVal sFun As String) As String

        Dim sReturn As String = ""                      ' return string
        sReturn = funGetExpression()                    ' get expression 2
        sReturn = funGetExpression() + sFun + sReturn   ' build expr1+op.+expr2

        Return sReturn                                  ' return complete string

    End Function

    '' OK-Function funBaseTokenD9, List Token and matrix token
    '  returns list as string
    Private Function funBaseTokenD9(ByVal sFun As String) As String

        Dim sReturn As String = "{"         ' start with open bracket
        Dim sComma As String = ""           ' used to add comma between arguments

        If byteProgramLine(intLineArrayPointer) = &HD9 Then boolMatrixMode = True ' ' if matrix token
        If boolMatrixMode = True Then sReturn = "[" '                               ' change bracket

        While byteProgramLine(intLineArrayPointer) <> &HE5                          ' until list end
            If byteProgramLine(intLineArrayPointer) <> &HD9 Then sReturn += sComma '' add comma if needed
            sReturn += funGetExpression()                                           ' get next expression
            sComma = ","                                                            ' enable comma
        End While                                                                   '
        intLineArrayPointer += 1                                                    ' inc past end tag

        If boolMatrixMode = True Then       ' if in matrix mode
            sReturn += "]"                  ' use this bracket
        Else : sReturn += "}" '             ' otherwise use this
        End If                              '

        Return sReturn                      ' return matrix or list as string

    End Function

    '' Function funBaseTokenDA, execute function/program
    '  returns string, format: sName(expr1,expr2)
    Private Function funBaseTokenDA(ByVal sFun As String) As String

        Dim sReturn As String   ' return string
        sReturn = funGetExpression() + funVarParameter("")
        Return sFun + sReturn

    End Function

    '' Function funBaseTokenDC, gets argument list from
    '  begginng of tokenized programs
    Private Function funBaseTokenDC(ByVal sFun As String) As String

        Dim sReturn As String           ' return string
        intLineArrayPointer += 3        ' bypass flags

        sReturn = sVariableName + funVarParameter("") + vbCrLf

        Return sReturn

    End Function

    '' OK-Function funBaseTokenE3, *function token, sets intTokenMode = 2
    '  and sets open "(" mode on
    Private Function funBaseTokenE3(ByVal sFun As String) As String

        intTokenMode = 2        ' sets token mode to 2
        Return sFun + ""        ' return token string

    End Function

    '' OK-Function funBaseTokenE4, *program token, sets intTokenMode = 3
    '  and returns ""
    Private Function funBaseTokenE4(ByVal sFun As String) As String

        intTokenMode = 3        ' set token mode to 3
        Return sFun + ""        ' returns token string

    End Function

    '' Function funBaseTokenE5, 'End_Tag' , closes out ")" and "}" for end
    '  ***** may need to modify for :'s in line, returns closing
    Private Function funBaseTokenE5(ByVal sFun As String) As String

        Return ""       ' returns nothing

    End Function

    '' OK-Function funBaseTokenE8, handles E7 (seperator token) E8 EOL token
    '  returns end of line string
    Private Function funBaseTokenE8(ByVal sFun As String) As String

        Dim sReturn As String = ""                  ' return string
        sReturn = sLineSuffix                       ' adds line suffix (if any)
        sLineSuffix = ""                            ' clear line suffix
        intLineArrayPointer += 1                    ' bypass &H00 following &HE7 and &HE8
        Return sReturn + sFun                       ' returns string

    End Function

    '' OK-Function funBaseTokenF0, signifies VAR name string to follow
    '  sets flag to signify VAR name to follow, *****not implemented, returns ""
    Private Function funBaseTokenF0(ByVal sFun As String) As String

        boolIsVarString = True      ' sets is variable string flag
        Return sFun + ""            ' returns token string

    End Function

#End Region


#Region "FuncToken Functions"

    '' Function funToken0F setgraph
    '  may need to change to variable paramater function
    Private Function funFunctionToken0F(ByVal sFun As String) As String

        Return sFun + "(" + funGetExpression() + "," + funGetExpression() + ")"

    End Function


    '' Function funFunctionToken0A, single expression function
    '  returns string, format: function(expr1)
    Private Function funFunctionToken0A(ByVal sFun As String) As String

        Return sFun + "(" + funGetExpression() + ")"

    End Function

    '' Function funFunctionToken14, single expression function
    '  returns string, format: function(expr1)
    Private Function funFunctionToken14(ByVal sFun As String) As String
        Dim sReturn As String
        sReturn = funGetExpression()
        Return funGetExpression() + sFun + sReturn

    End Function

#End Region


#Region "ProgramToken Functions"

    '' OK-Function funProgramToken3B, 'If EXPR Then', adds " Then" suffix
    '  returns ""
    Private Function funProgramToken3B(ByVal sFun As String) As String

        sLineSuffix = " Then"       ' set suffix
        Return sFun

    End Function

    '' OK-Function funProgramToken15, EndWhile
    '  Removes displacement bytes
    Private Function funProgramToken15(ByVal sFun As String) As String

        While byteProgramLine(intLineArrayPointer) <> 232       ' run till end of line
            intLineArrayPointer += 2                            ' inc array pointer
        End While

        Return sFun                                             ' return nothing
    End Function

    '' OK-Function funProgramToken86, Define Expr1 = Expr2
    '  Removes displacement bytes
    Private Function funProgramToken86(ByVal sFun As String) As String

        Dim sReturn As String                                   ' return string
        sReturn = funGetExpression() + "=" + funGetExpression() ' inc array pointer
        Return sFun + sReturn                                   ' return nothing

    End Function


#End Region


#Region "Package to Save"

    '' Function funParseText
    '  Parses raw program text file and converts it to text form for saving
    Private Function funParseText(ByRef sProgramText As String) As String

        Dim sTemp As String                         ' Temp string storage
        Dim intVarLength As Integer                 ' Length of variable name to remove
        Dim sVarName As String = Me.variableName    ' Holds variable name
        Dim intI As Integer                         ' used for loops
        Dim intL As Integer                         ' used to hold string lengths
        Dim charArray() As Char                   ' holds strings converted to arrays

        intVarLength = sVarName.Length              ' holds length of variable name
        If sProgramText.StartsWith(sVarName) Then   ' if the variable name is at the beginning
            sProgramText = sProgramText.Remove(0, intVarLength)     ' remove it
        End If                                                      '

        sProgramText = sProgramText.Replace(vbCrLf, Chr(&HD))       ' swap CrLf with &H0D
        sProgramText = sProgramText.Replace(Chr(&HBB), Chr(&H16))   ' swap &HBB with &H16
        sProgramText = sProgramText.Replace(Chr(&HAA), Chr(&HAD))   ' swap &HAA with &HAD

        ReDim byteRawProgram(sProgramText.Length + 99)             ' resize array to hold program

        '' saves calculator name into array
        '
        charArray = sCalcName.ToCharArray             ' convert calc name to char array
        For intI = 0 To 7                                           ' store calculator name
            byteRawProgram(intI) = CByte(AscW(charArray(intI)))     ' into program array
        Next                                                        '

        '' saves seperator 1 into array
        '
        For intI = 8 To 9                                           ' store seperator 1
            byteRawProgram(intI) = byteSeperator1(intI - 8)         ' into program array
        Next                                                        '

        '' save folder name into array
        '
        intL = sFolderName.Length - 1                               ' get length of folder name
        charArray = sFolderName.ToCharArray                         ' convert folder name to char array
        For intI = &HA To &HA + intL                                ' store folder name
            byteRawProgram(intI) = CByte(AscW(charArray(intI - &HA)))   ' into program array
        Next                                                        '

        '' save comment text
        '
        intL = sProgComment.Length - 1
        charArray = sProgComment.ToCharArray                        ' convert comment text to char array
        For intI = &H12 To &H12 + intL                              ' store comment text
            byteRawProgram(intI) = CByte(AscW(charArray(intI - &H12)))  ' into program array
        Next                                                        '

        '' save seperator 2
        '
        For intI = &H3B To &H3F                                     ' store seperator 2
            byteRawProgram(intI) = byteSeperator2(intI - &H3B)      ' into program array
        Next                                                        '

        '' save variable name
        '
        intL = sVariableName.Length - 1                             ' get length of variable name
        charArray = sVariableName.ToCharArray                       ' convert variable name to char array
        For intI = &H40 To &H40 + intL                              ' store comment text
            byteRawProgram(intI) = CByte(AscW(charArray(intI - &H40)))  ' into program array
        Next                                                        '

        '' save file type
        '
        byteRawProgram(&H48) = byteFileType                         ' store file type to array

        '' save seperator 3
        '
        For intI = &H49 To &H4B                                     ' store seperator 3
            byteRawProgram(intI) = byteSeperator3(intI - &H49)      ' into program array
        Next                                                        '

        ''  save file length
        '
        Dim sLength As String                                       ' holds hex length string
        intFileLength = sProgramText.Length + 100                   ' set file length
        sLength = Hex(intFileLength)                                ' convert to hex
        sLength = sLength.PadRight(8, CChar("0"))                   ' pad to 4 bytes long
        For intI = &H4C To &H4F                                     ' store file length
            byteRawProgram(intI) = CByte("&H" + sLength.Substring(((intI - &H4C) * 2), 2)) '
        Next                                                        '

        '' save seperator 4
        '
        For intI = &H50 To &H55                                     ' store seperator 4
            byteRawProgram(intI) = byteSeperator4(intI - &H50)      ' into program array
        Next                                                        '

        '' save program length
        '
        intProgramSize = sProgramText.Length + 10                   ' set program length
        sLength = Hex(intProgramSize)                               ' convert to hex
        sLength = sLength.PadRight(4, CChar("0"))                   ' pad to 2 bytes long
        For intI = &H56 To &H57                                     ' store file length
            byteRawProgram(&H56) = CByte("&H" + sLength.Substring(2, 2))
            byteRawProgram(&H57) = CByte("&H" + sLength.Substring(0, 2)) '
        Next

        '' save program text
        '
        intL = sProgramText.Length - 1                              ' get length of program text
        charArray = sProgramText.ToCharArray                        ' convert program text to char array
        For intI = &H58 To &H58 + intL                              ' store program text
            byteRawProgram(intI) = CByte(AscW(charArray(intI - &H58)))  ' into program array
        Next

        ''  save seperator 5
        '
        intL = intI + 1                                             ' start just past program
        For intI = intL To intL + 1                                 ' store seperator 5
            byteRawProgram(intI) = byteSeperator5(intI - intL)      ' into program array
        Next                                                        '

        '' save file type
        '
        If byteFileType = &H12 Then
            byteRawProgram(intL + 2) = byteProgTag(0)      ' into program array
            byteRawProgram(intL + 3) = byteProgTag(1)      ' into program array
        End If
        If byteFileType = &H13 Then
            byteRawProgram(intL + 2) = byteFuncTag(0)      ' into program array
            byteRawProgram(intL + 3) = byteFuncTag(1)      ' into program array
        End If

        '' save end tag
        '
        byteRawProgram(intL + 4) = byteEndTag

        '' save seperator 6
        For intI = intL + 5 To intL + 7                                 ' store seperator 6
            byteRawProgram(intI) = byteSeperator6(intI - intL - 5)      ' into program array
        Next

        '' save end mark
        '
        byteRawProgram(intL + 8) = byteEndMark

        '' save check sum
        Dim intT As Integer
        For intI = 86 To byteRawProgram.Length - 5
            intT += byteRawProgram(intI)
        Next
        sLength = Hex(intT)                                         ' convert to hex
        If sLength.Length < 4 Then sLength = sLength.PadLeft(4, CChar("0")) ' pad to 2 bytes long
        If sLength.Length > 4 Then sLength.Substring(0, 3)
        byteRawProgram(intL + 9) = CByte("&H" + sLength.Substring(2, 2))
        byteRawProgram(intL + 10) = CByte("&H" + sLength.Substring(0, 2))



        Return ""

    End Function

#End Region


#Region "New File"


    Private Sub subNewFile(ByVal byteType As Byte)

    End Sub


#End Region


#End Region


End Class
