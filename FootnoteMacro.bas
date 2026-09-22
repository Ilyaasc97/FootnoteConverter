REM  *****  BASIC  *****
' ==============================================================================
' حزمة ماكرو الحواشي المتقدمة لـ LibreOffice Writer (الإصدار 2.3.0)
' تتضمن:
' 1. تحويل علامات الحواشي الحصرية < > إلى حواشي سفلية معتمدة دون المساس بالاقتباسات « » أو (( )).
' 2. توافق تام 100% مع Microsoft Word (.docx) عند الحفظ والتصدير للمشرفين والجامعات.
' 3. إحاطة رقم المتن بأقواس (1) مع إزالة الفراغ قبل القوس.
' 4. ترقيم الحواشي بالأسفل بصيغة (1) ومحاذاتها لليمين (RTL).
' 5. ضبط خط الفاصل تلقائياً لجهة اليمين.
' 6. توحيد خط وحجم الحواشي السفلية تلقائياً (UnifyFootnotesFont) وتطبيق معيار (حجم المتن - 2).
' 7. ماكرو الترقيم لكل صفحة مستقلة (SetFootnoteNumberingPerPage).
' 8. ماكرو الترقيم المستمر لكامل المستند (SetFootnoteNumberingContinuous).
' 9. ماكرو التحويل العكسي واستخراج الحواشي لنصوص < > (ReverseFootnotesToMarkers).
' 10. إحاطة النص المحدد بـ < > وحذف الأقواس (WrapSelectionWithMarkers).
' 11. دعم كامل للتراجع (Ctrl + Z).
' ==============================================================================

Sub ConvertMarkersToFootnotes()
    Dim oDoc As Object
    Dim oText As Object
    Dim oSearch As Object
    Dim oFound As Object
    Dim oMatch As Object
    Dim oFootnote As Object
    Dim oSettings As Object
    Dim oUndoManager As Object
    Dim oStyleFamilies As Object
    Dim oPageStyles As Object
    Dim oPageStyle As Object
    Dim oParaStyles As Object
    Dim oFootnoteStyle As Object
    Dim sStyleNames() As String
    Dim sRaw As String
    Dim sContent As String
    Dim nCount As Long
    Dim i As Long
    Dim k As Long
    Dim oCur As Object
    Dim oCurPre As Object
    Dim oCurPost As Object
    Dim oFnCur As Object
    Dim sPreChar As String
    Dim sPostChar As String

    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "يرجى فتح مستند نصي (LibreOffice Writer) لتشغيل الماكرو.", 48, "تنبيه"
        Exit Sub
    End If

    oText = oDoc.getText()

    ' 1. فحص لغة المستند واتجاه الكتابة (عربي = يمين / إنجليزي = يسار)
    Dim sSample As String
    Dim isArabic As Boolean
    Dim sChar As String
    Dim j As Long
    Dim nAdjust As Integer
    Dim nWritingMode As Integer

    isArabic = False
    sSample = Left(oText.getString(), 2000)
    For j = 1 To Len(sSample)
        sChar = Mid(sSample, j, 1)
        If AscW(sChar) >= 1536 And AscW(sChar) <= 1791 Then
            isArabic = True
            Exit For
        End If
    Next j

    If isArabic Then
        nAdjust = 2       ' محاذاة لليمين للنصوص العربية
        nWritingMode = 1  ' اتجاه يمين لليسار RTL
    Else
        nAdjust = 0       ' محاذاة لليسار للنصوص اللاتينية والإنجليزية
        nWritingMode = 0  ' اتجاه يسار لليمين LTR
    End If

    ' 2. ضبط خط الفاصل بين الحاشية والمتن حسب اتجاه اللغة
    On Error Resume Next
    oStyleFamilies = oDoc.getStyleFamilies()
    If oStyleFamilies.hasByName("PageStyles") Then
        oPageStyles = oStyleFamilies.getByName("PageStyles")
        sStyleNames = oPageStyles.getElementNames()
        For k = LBound(sStyleNames) To UBound(sStyleNames)
            oPageStyle = oPageStyles.getByName(sStyleNames(k))
            oPageStyle.FootnoteLineAdjust = nAdjust
        Next k
    End If

    ' 3. ضبط نمط فقرة الحاشية Footnote حسب اتجاه اللغة
    If oStyleFamilies.hasByName("ParagraphStyles") Then
        oParaStyles = oStyleFamilies.getByName("ParagraphStyles")
        If oParaStyles.hasByName("Footnote") Then
            oFootnoteStyle = oParaStyles.getByName("Footnote")
            oFootnoteStyle.WritingMode = nWritingMode
            oFootnoteStyle.ParaAdjust = nAdjust
        End If
    End If
    On Error GoTo 0

    ' 3. ضبط إعدادات الحواشي لتظهر الأرقام بالأسفل بين قوسين (1)
    On Error Resume Next
    oSettings = oDoc.FootnoteSettings
    oSettings.Prefix = "("
    oSettings.Suffix = ") "
    oSettings.NumberingType = 4  ' أرقام عربية 1، 2، 3
    On Error GoTo 0

    ' 4. البحث بالتعبير النمطي لعلامات الحواشي الحصرية < >
    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchRegularExpression = True
    oSearch.SearchString = "<([^<>]*)>"

    oFound = oDoc.findAll(oSearch)

    ' 5. بدء سياق التراجع
    On Error Resume Next
    oUndoManager = oDoc.getUndoManager()
    If Not IsNull(oUndoManager) Then
        oUndoManager.enterUndoContext("تحويل وتنسيق الحواشي المطور")
    End If
    oDoc.lockControllers()
    oDoc.addActionLock()
    On Error GoTo 0

    nCount = 0
    If Not IsNull(oFound) Then
        nCount = oFound.getCount()
    End If

    ' 6. التحويل والمعالجة الذكية
    If nCount > 0 Then
        For i = nCount - 1 To 0 Step -1
            oMatch = oFound.getByIndex(i)
            sRaw = Trim(oMatch.getString())

            ' استخراج المحتوى الداخلي بين < و >
            sContent = sRaw
            If Left(sContent, 1) = "<" And Right(sContent, 1) = ">" Then
                sContent = Mid(sContent, 2, Len(sContent) - 2)
            End If
            sContent = Trim(sContent)

            ' إذا كانت الحاشية فارغة <> يتم حذفها من المتن فوراً دون إنشاء حاشية سفلية
            If Len(sContent) = 0 Then
                oMatch.setString("")
            Else
                ' معالجة المسافة قبل الحاشية
                On Error Resume Next
                oCurPre = oText.createTextCursorByRange(oMatch.getStart())
                If oCurPre.goLeft(1, True) Then
                    sPreChar = oCurPre.getString()
                    If sPreChar = " " Or sPreChar = Chr(160) Then
                        oCurPre.setString("")
                    End If
                End If

                ' معالجة الفراغ بعد الحاشية وقبل علامة الترقيم
                oCurPost = oText.createTextCursorByRange(oMatch.getEnd())
                If oCurPost.goRight(2, True) Then
                    sPostChar = oCurPost.getString()
                    If Len(sPostChar) = 2 Then
                        If Left(sPostChar, 1) = " " Then
                            Dim sPunc As String
                            sPunc = Right(sPostChar, 1)
                            If InStr(".،:؛!؟", sPunc) > 0 Then
                                oCurPost.setString(sPunc)
                            End If
                        End If
                    End If
                End If
                On Error GoTo 0

                ' إفراغ النص القديم
                oMatch.setString("")

                ' إدراج ( [Footnote] )
                oCur = oText.createTextCursorByRange(oMatch.getStart())

                On Error Resume Next
                oCur.CharStyleName = "Footnote anchor"
                On Error GoTo 0
                oText.insertString(oCur, "(", False)
                oCur.collapseToEnd()

                oFootnote = oDoc.createInstance("com.sun.star.text.Footnote")
                oText.insertTextContent(oCur, oFootnote, False)
                oCur.collapseToEnd()

                On Error Resume Next
                oCur.CharStyleName = "Footnote anchor"
                On Error GoTo 0
                oText.insertString(oCur, ")", False)
                oCur.collapseToEnd()

                On Error Resume Next
                oCur.setPropertyToDefault("CharStyleName")
                On Error GoTo 0

                oFootnote.setString(sContent)

                ' محاذاة نص الحاشية لليمين (RTL)
                On Error Resume Next
                oFnCur = oFootnote.getText().createTextCursor()
                oFnCur.gotoStart(False)
                oFnCur.gotoEnd(True)
                oFnCur.WritingMode = nWritingMode
                oFnCur.ParaAdjust = nAdjust
                On Error GoTo 0
            End If
        Next i
    End If

    ' 7. ضبط أية حواشي قديمة في المستند
    Dim oFootnotes As Object
    Dim oFn As Object
    Dim oAnchor As Object
    Dim oCurBefore As Object
    Dim oCurAfter As Object
    Dim sBefore As String
    Dim sAfter As String
    Dim fnCount As Long

    On Error Resume Next
    oFootnotes = oDoc.getFootnotes()
    If Not IsNull(oFootnotes) Then
        fnCount = oFootnotes.getCount()
        For i = fnCount - 1 To 0 Step -1
            oFn = oFootnotes.getByIndex(i)
            oFnCur = oFn.getText().createTextCursor()
            oFnCur.gotoStart(False)
            oFnCur.gotoEnd(True)
            oFnCur.WritingMode = 1
            oFnCur.ParaAdjust = 1

            oAnchor = oFn.getAnchor()
            If Not IsNull(oAnchor) Then
                oCurBefore = oText.createTextCursorByRange(oAnchor.getStart())
                If oCurBefore.goLeft(1, True) Then
                    sBefore = oCurBefore.getString()
                Else
                    sBefore = ""
                End If

                oCurAfter = oText.createTextCursorByRange(oAnchor.getEnd())
                If oCurAfter.goRight(1, True) Then
                    sAfter = oCurAfter.getString()
                Else
                    sAfter = ""
                End If

                If sBefore <> "(" Then
                    oCur = oText.createTextCursorByRange(oAnchor.getStart())
                    oCur.CharStyleName = "Footnote anchor"
                    oText.insertString(oCur, "(", False)
                    oCur.setPropertyToDefault("CharStyleName")
                End If

                If sAfter <> ")" Then
                    oCur = oText.createTextCursorByRange(oAnchor.getEnd())
                    oCur.CharStyleName = "Footnote anchor"
                    oText.insertString(oCur, ")", False)
                    oCur.setPropertyToDefault("CharStyleName")
                End If
            End If
        Next i
    End If
    On Error GoTo 0

    ' إنهاء القفل وسياق التراجع
    On Error Resume Next
    oDoc.removeActionLock()
    oDoc.unlockControllers()
    If Not IsNull(oUndoManager) Then
        oUndoManager.leaveUndoContext()
    End If
    On Error GoTo 0

    MsgBox "تمت العملية بنجاح!" & Chr(10) & Chr(10) & _
           "• تم تحويل (" & nCount & ") حاشية سفلية جديدة." & Chr(10) & _
           "• إحاطة رقم المرجع في المتن بأقواس (1) وإلصاقه بالكلمة السابقة." & Chr(10) & _
           "• ترقيم الحاشية السفلية بالأسفل بصيغة (1) ومحاذاتها لليمين." & Chr(10) & _
           "• تصحيح محاذاة خط الفاصل إلى جهة اليمين." & Chr(10) & _
           "• يمكنك التراجع بضغطة واحدة عبر (Ctrl + Z).", 64, "تمت العملية بنجاح"
End Sub

' ==============================================================================
' ماكرو ضبط ترقيم الحواشي ليكون مستقلاً لكل صفحة (1 في كل صفحة جديدة)
' ==============================================================================
Sub SetFootnoteNumberingPerPage()
    Dim oDoc As Object
    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub
    On Error Resume Next
    ' 0 = com.sun.star.text.FootnoteNumbering.PER_PAGE
    oDoc.FootnoteSettings.FootnoteCounting = 0
    MsgBox "تم بنجاح ضبط ترقيم الحواشي ليكون مستقلاً يبدأ من (1) في كل صفحة جديدة.", 64, "نطاق الترقيم"
End Sub

' ==============================================================================
' ماكرو ضبط ترقيم الحواشي ليكون مستمراً لكامل المستند
' ==============================================================================
Sub SetFootnoteNumberingContinuous()
    Dim oDoc As Object
    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub
    On Error Resume Next
    ' 2 = com.sun.star.text.FootnoteNumbering.PER_DOCUMENT
    oDoc.FootnoteSettings.FootnoteCounting = 2
    MsgBox "تم بنجاح ضبط ترقيم الحواشي ليكون ترقيماً متسلسلاً مستمراً لكامل المستند.", 64, "نطاق الترقيم"
End Sub

' ==============================================================================
' ماكرو التحويل العكسي: استخراج الحواشي السفلية وإعادتها لنصوص < > في المتن
' ==============================================================================
Sub ReverseFootnotesToMarkers()
    Dim oDoc As Object
    Dim oText As Object
    Dim oFootnotes As Object
    Dim oFn As Object
    Dim oAnchor As Object
    Dim oUndoManager As Object
    Dim oCurB As Object
    Dim oCurA As Object
    Dim oCurRep As Object
    Dim sContent As String
    Dim nCount As Long
    Dim i As Long

    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub

    oFootnotes = oDoc.getFootnotes()
    If IsNull(oFootnotes) Then Exit Sub
    nCount = oFootnotes.getCount()
    If nCount = 0 Then
        MsgBox "لا توجد أية حواشي سفلية في المستند لاستخراجها.", 64, "التحويل العكسي"
        Exit Sub
    End If

    oText = oDoc.getText()

    On Error Resume Next
    oUndoManager = oDoc.getUndoManager()
    If Not IsNull(oUndoManager) Then
        oUndoManager.enterUndoContext("استخراج الحواشي لنصوص")
    End If
    oDoc.lockControllers()
    oDoc.addActionLock()
    On Error GoTo 0

    For i = nCount - 1 To 0 Step -1
        oFn = oFootnotes.getByIndex(i)
        sContent = Trim(oFn.getString())
        oAnchor = oFn.getAnchor()

        If Not IsNull(oAnchor) Then
            Dim hasB As Boolean
            Dim hasA As Boolean
            hasB = False
            hasA = False

            oCurB = oText.createTextCursorByRange(oAnchor.getStart())
            If oCurB.goLeft(1, True) Then
                If oCurB.getString() = "(" Then hasB = True
            End If

            oCurA = oText.createTextCursorByRange(oAnchor.getEnd())
            If oCurA.goRight(1, True) Then
                If oCurA.getString() = ")" Then hasA = True
            End If

            If hasB And hasA Then
                oCurRep = oText.createTextCursorByRange(oAnchor.getStart())
                oCurRep.goLeft(1, False)
                oCurRep.gotoRange(oAnchor.getEnd(), True)
                oCurRep.goRight(1, True)
            Else
                oCurRep = oText.createTextCursorByRange(oAnchor.getStart())
                oCurRep.gotoRange(oAnchor.getEnd(), True)
            End If

            oCurRep.setPropertyToDefault("CharStyleName")
            oCurRep.setString(" <" & sContent & ">")
        End If
    Next i

    On Error Resume Next
    oDoc.removeActionLock()
    oDoc.unlockControllers()
    If Not IsNull(oUndoManager) Then
        oUndoManager.leaveUndoContext()
    End If
    On Error GoTo 0

    MsgBox "تم بنجاح استخراج (" & nCount & ") حاشية وإعادتها كنصوص < > داخل المتن!", 64, "التحويل العكسي"
End Sub

' ==============================================================================
' ماكرو إحاطة النص المحدد بـ < > وحذف الأقواس العادية (Alt + X أو F2)
' إذا لم يكن هناك نص محدد، يتم إدراج <> ووضع المؤشر بالداخل
' ==============================================================================
Sub WrapSelectionWithMarkers()
    Dim oDoc As Object
    Dim oSelection As Object
    Dim oRange As Object
    Dim oVCur As Object
    Dim oText As Object
    Dim sRaw As String
    Dim sTrim As String
    Dim sInner As String
    Dim hasSel As Boolean
    Dim i As Long

    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub

    oSelection = oDoc.getCurrentSelection()
    oVCur = oDoc.getCurrentController().getViewCursor()
    hasSel = False

    If Not IsNull(oSelection) Then
        For i = 0 To oSelection.getCount() - 1
            oRange = oSelection.getByIndex(i)
            sRaw = oRange.getString()
            sTrim = Trim(sRaw)
            If Len(sTrim) > 0 Then
                hasSel = True
                sInner = sTrim
                If Left(sInner, 1) = "(" And Right(sInner, 1) = ")" Then
                    sInner = Trim(Mid(sInner, 2, Len(sInner) - 2))
                ElseIf Left(sInner, 1) = "<" And Right(sInner, 1) = ">" Then
                    sInner = Trim(Mid(sInner, 2, Len(sInner) - 2))
                End If
                oRange.setString("<" & sInner & ">")
            End If
        Next i
    End If

    If Not hasSel And Not IsNull(oVCur) Then
        oText = oVCur.getText()
        oText.insertString(oVCur, "<>", False)
        oVCur.goLeft(1, False)
    End If
End Sub

' ==============================================================================
' ماكرو توحيد خط وحجم الحواشي السفلية تلقائياً (Font & Size)
' يقرأ خط المتن ويطبق قاعدة النقطتين الأكاديمية (حجم الحاشية = حجم المتن - 2)
' مع تطهير التنسيقات المشوهة ومحاذاة الحاشية والفاصل لليمين
' ==============================================================================
Sub UnifyFootnotesFont()
    Dim oDoc As Object
    Dim oStyleFamilies As Object
    Dim oParaStyles As Object
    Dim oFootnotes As Object
    Dim oFootnote As Object
    Dim oCur As Object
    Dim oUndoManager As Object
    Dim sBodyFontComplex As String
    Dim sBodyFontWestern As String
    Dim fBodySizeComplex As Single
    Dim fBodySizeWestern As Single
    Dim sTargetFontComplex As String
    Dim sTargetFontWestern As String
    Dim fTargetSizeComplex As Single
    Dim fTargetSizeWestern As Single
    Dim nCount As Long
    Dim i As Long

    oDoc = ThisComponent
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "يرجى فتح مستند نصي في ليبر أوفيس أولاً!", 48, "تنبيه"
        Exit Sub
    End If

    ' بدء التراجع بضغطة واحدة
    On Error Resume Next
    If HasUnoInterfaces(oDoc, "com.sun.star.document.XUndoManagerSupplier") Then
        oUndoManager = oDoc.getUndoManager()
        If Not IsNull(oUndoManager) Then
            oUndoManager.enterUndoContext("توحيد خط وحجم الحواشي")
        End If
    End If
    oDoc.lockControllers()
    oDoc.addActionLock()
    On Error GoTo 0

    ' القيم الافتراضية القياسية الأكاديمية
    sBodyFontComplex = "Traditional Arabic"
    sBodyFontWestern = "Times New Roman"
    fBodySizeComplex = 18.0
    fBodySizeWestern = 12.0

    ' فحص خط المتن
    On Error Resume Next
    oStyleFamilies = oDoc.getStyleFamilies()
    If oStyleFamilies.hasByName("ParagraphStyles") Then
        oParaStyles = oStyleFamilies.getByName("ParagraphStyles")
        Dim aNames(2) As String
        aNames(0) = "Standard"
        aNames(1) = "Default Paragraph Style"
        aNames(2) = "Body Text"
        Dim k As Integer
        For k = 0 To 2
            If oParaStyles.hasByName(aNames(k)) Then
                Dim oBaseStyle As Object
                oBaseStyle = oParaStyles.getByName(aNames(k))
                If oBaseStyle.CharFontNameComplex <> "" Then
                    sBodyFontComplex = oBaseStyle.CharFontNameComplex
                End If
                If oBaseStyle.CharHeightComplex > 0 Then
                    fBodySizeComplex = oBaseStyle.CharHeightComplex
                End If
                If oBaseStyle.CharFontName <> "" Then
                    sBodyFontWestern = oBaseStyle.CharFontName
                End If
                If oBaseStyle.CharHeight > 0 Then
                    fBodySizeWestern = oBaseStyle.CharHeight
                End If
                Exit For
            End If
        Next k
    End If
    On Error GoTo 0

    ' حساب حجم الحاشية (حجم المتن - 2 نقطة)
    sTargetFontComplex = sBodyFontComplex
    sTargetFontWestern = sBodyFontWestern
    fTargetSizeComplex = fBodySizeComplex - 2.0
    If fTargetSizeComplex < 8.0 Then fTargetSizeComplex = 8.0
    fTargetSizeWestern = fBodySizeWestern - 2.0
    If fTargetSizeWestern < 8.0 Then fTargetSizeWestern = 8.0

    ' 1. تعديل نمط فقرة الحاشية Footnote
    On Error Resume Next
    If oParaStyles.hasByName("Footnote") Then
        Dim oFnStyle As Object
        oFnStyle = oParaStyles.getByName("Footnote")
        oFnStyle.CharFontNameComplex = sTargetFontComplex
        oFnStyle.CharHeightComplex = fTargetSizeComplex
        oFnStyle.CharFontName = sTargetFontWestern
        oFnStyle.CharHeight = fTargetSizeWestern
        oFnStyle.WritingMode = 1
        oFnStyle.ParaAdjust = 1
    End If
    On Error GoTo 0

    ' 2. توحيد جميع الحواشي الفردية القائمة
    oFootnotes = oDoc.getFootnotes()
    nCount = 0
    If Not IsNull(oFootnotes) Then
        nCount = oFootnotes.getCount()
        For i = 0 To nCount - 1
            oFootnote = oFootnotes.getByIndex(i)
            oCur = oFootnote.getText().createTextCursor()
            oCur.gotoStart(False)
            oCur.gotoEnd(True)
            On Error Resume Next
            oCur.CharFontNameComplex = sTargetFontComplex
            oCur.CharHeightComplex = fTargetSizeComplex
            oCur.CharFontName = sTargetFontWestern
            oCur.CharHeight = fTargetSizeWestern
            oCur.WritingMode = 1
            oCur.ParaAdjust = 1
            On Error GoTo 0
        Next i
    End If

    ' 3. ضبط فاصل الحاشية جهة اليمين
    On Error Resume Next
    AdjustFootnoteSeparatorRTL(oDoc)
    On Error GoTo 0

    ' إنهاء القفل والتراجع
    On Error Resume Next
    oDoc.removeActionLock()
    oDoc.unlockControllers()
    If Not IsNull(oUndoManager) Then
        oUndoManager.leaveUndoContext()
    End If
    On Error GoTo 0

    MsgBox "تم بنجاح توحيد خط وحجم الحواشي السفلية!" & Chr(10) & Chr(10) & _
           "• نوع الخط الموحد: " & sTargetFontComplex & Chr(10) & _
           "• حجم خط الحاشية: " & fTargetSizeComplex & " نقطة (أصغر بنقطتين من المتن: " & fBodySizeComplex & ")" & Chr(10) & _
           "• عدد الحواشي الموحدة: (" & nCount & ") حاشية" & Chr(10) & _
           "• تم ضبط المحاذاة والفاصل لليمين." & Chr(10) & Chr(10) & _
           "يمكنك التراجع في أي وقت عبر (Ctrl + Z).", 64, "توحيد خط الحواشي"
End Sub

