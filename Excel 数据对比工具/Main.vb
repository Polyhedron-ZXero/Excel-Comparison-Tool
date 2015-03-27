'陈乔异 (Qiaoyi Chen)
'联想（上海）有限公司

'更新日志：
'
'版本：1.2.0.0
'日期：2014/08/06
'----------------
'   1. 增加一个线程，使数据对比时窗口保持活动
'   2. 增加运行时的进度条
'   3. 用 AndAlso 和 OrElse 替换 And 和 Or
'   4. 修正高 DPI 下文字显示不完整的问题
'   5. 修改程序图标
'   6. 增加更多的异常处理
'
'版本：1.1.0.1
'日期：2014/07/25
'----------------
'   1. 增加界面透明度调整滚动条
'   2. 调整控件的 Tab 键控制顺序
'
'版本：1.1.0.0
'日期：2014/07/22
'----------------
'   1. 添加多语言支持与实时切换（简体中文、繁体中文、英语、日语）
'
'版本：1.0.0.2
'日期：2014/07/16
'----------------
'   1. 添加在未安装 Excel 时的错误提示
'   2. 解决保存后再对比时提示需要重新打开 Excel 文件的 bug
'
'版本：1.0.0.1
'日期：2014/07/16
'----------------
'   1. 对比字符串去除了前后空格
'   2. 打开文件前生成一个空 Excel 进程，解决用户手动打开其他 Excel 文件时造成临时副本同时打开的问题
'   3. 出现严重错误并退出时会自动终止所有 EXCEL.EXE 进程，防止文件被锁定
'
'版本：1.0.0.0
'日期：2014/07/15
'----------------
'   1. 确定基本界面和基本操作方式
'   2. 优化搜索效率（使用数组），搜索完成后显示运行时间
'   3. 增加基本错误处理和异常处理机制
'   4. 增加文本框数据合法性检测机制
'   5. 确定对单元格标注颜色的定义：互相匹配 = 绿（背景）；无匹配 = 红（背景）；自身重复 = 黄（字体）
'   6. 调整文件打开方式（创建原文件的临时副本，并在临时副本上操作）
'   7. 增加全局布尔变量，在用户关闭程序或重新打开文件时，如果有未保存的内容则提示用户

Imports Microsoft.Office.Interop

'sheet1 = book1.Worksheets("Sheet1")  设置当前工作表
'sheet1.Range("A1").Value  获取单元格数值
'sheet1.Range("A1").Interior.ColorIndex = 4  设置单元格背景色（4 = 绿，6 = 黄，3 = 红）
'sheet1.Range("A1").Resize(100, 3).Value = dataArray  把二维数组的数据直接传输到从 A1 开始的 100 行 * 3 列

'MSDN 参考资料：
'http://support.microsoft.com/kb/306022

Public Class Main
    Dim file1 As String, file2 As String  '原文件名（完整路径）
    Const tempFile1 As String = "C:\Users\Public\~$xls_tmp1", tempFile2 As String = "C:\Users\Public\~$xls_tmp2"  '临时副本文件名（完整路径）
    Dim activeSheet1 As Integer, activeSheet2 As Integer  '活动工作表的 index
    Dim isNeedSave As Boolean  '是否需要保存
    Dim isOpened As Boolean  '文件是否处于打开状态

    '关于 emptyExcel：
    '   Excel 在打开一个文件时会检查进程中是否已存在一个 Application (EXCEL.EXE)，
    '   如果存在，则直接使用第一个 Application，而不会再打开一个进程。
    '   正在操作中的临时副本所在的 Application 被设置为 Visible = False，应该是不可见的。
    '   但是如果运行时用户手动打开另一个 Excel 文件，此文件则会在当前 Application 中被打开，
    '   且 Visible = False 会被重载而显示出临时副本，这会造成用户误关闭或误操作临时副本，
    '   可能导致出错。
    '   解决方法为先创建一个 emptyExcel 的 Application，如果用户另外打开其他
    '   Excel 文件，此文件只会使用 emptyExcel 的 Application 进程显示，而不会影响到
    '   临时副本文件的内部操作。
    Dim emptyExcel As Excel.Application, excel As Excel.Application
    Dim book1 As Excel.Workbook, book2 As Excel.Workbook
    Dim sheet1 As Excel.Worksheet, sheet2 As Excel.Worksheet

    Structure CompareArray
        Dim Data As String  '单元格内容
        Dim IsMatched As Boolean  '是否搜索到相同字符串
        Dim IsRepeated As Boolean  '是否重复项
    End Structure

    Enum Language
        zh_CN
        zh_TW
        en_US
        ja_JP
    End Enum

    '默认语言
    Dim lang As Language = Language.zh_CN

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
                MsgBox("程序已经在运行中。", MsgBoxStyle.Exclamation, "提示")
                Me.Dispose()
            End If
            isNeedSave = False
            isOpened = False
            file1 = Nothing
            file2 = Nothing
            emptyExcel = New Excel.Application
            excel = New Excel.Application
            emptyExcel.Visible = False
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False  '略提升速度
            book1 = Nothing
            book2 = Nothing

            '进度条 Visible 的初始属性必须设置为 True（启动后可以重设为 False），运行时才有动画过渡效果（使用多线程的情况下），并不会造成第二次运行时进度条无法显示
            ProgressBar1.Visible = False

            '解决无法跨线程操作控件的问题（错误“线程间操作无效：从不是创建控件...的线程访问它”）
            Control.CheckForIllegalCrossThreadCalls = False
        Catch ex As System.Runtime.InteropServices.COMException
            '注：未激活的 Excel 会在打开后弹出激活窗口，如果不把激活窗口关闭就用此软件打开 Excel 文档，程序会抛出异常并崩溃
            MsgBox("未检测到 Excel 软件！" & Chr(13) & Chr(13) & "此工具需要 Excel 支持，请确保计算机中已安装 Microsoft Excel 2007 或以上版本。" & Chr(13) & Chr(13) & "Excel is not detected!" & Chr(13) & Chr(13) & "This tool requires support of Excel. Please make sure that Microsoft Excel 2007 or higher version is installed on this computer.", MsgBoxStyle.Critical, "错误 Error")
            Me.Dispose()
        End Try
    End Sub

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If isNeedSave Then
            Select Case lang
                Case Language.zh_CN
                    If MsgBox("已修改的文档未保存，直接退出将导致数据丢失。" & Chr(13) & "是否放弃保存并退出程序？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "关闭程序").Equals(vbNo) Then
                        e.Cancel = True
                    End If
                Case Language.zh_TW
                    If MsgBox("已修改的檔案未保存，直接退出將導致數據丟失。" & Chr(13) & "是否放棄保存並退出程序？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "關閉程序").Equals(vbNo) Then
                        e.Cancel = True
                    End If
                Case Language.en_US
                    If MsgBox("The modified documents have not been saved. Exiting without saving will cause data loss." & Chr(13) & "Do you want to exit without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "Program Closing").Equals(vbNo) Then
                        e.Cancel = True
                    End If
                Case Language.ja_JP
                    If MsgBox("変更されたドキュメントは保存されていません。保存せずに終了すると、データが失われます。" & Chr(13) & "保存せずにプログラムを終了しますか。", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "プログラムの終了").Equals(vbNo) Then
                        e.Cancel = True
                    End If
            End Select
        End If
    End Sub

    Private Sub Main_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            book1.Close()
            book2.Close()
        Catch ex As Exception
        End Try
        System.IO.File.Delete(tempFile1)
        System.IO.File.Delete(tempFile2)
        excel.Quit()
        emptyExcel.Quit()
        GC.Collect()
    End Sub

    '浏览文件 1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            If InStr(OpenFileDialog1.FileName.Length - 4, OpenFileDialog1.FileName, ".xls", CompareMethod.Text) <> 0 Then
                file1 = OpenFileDialog1.FileName
                TextBox1.Clear()
                TextBox1.AppendText(file1)
            Else
                Select Case lang
                    Case Language.zh_CN
                        MsgBox("所选文件不是 Excel 文档。", MsgBoxStyle.Critical, "错误")
                    Case Language.zh_TW
                        MsgBox("所選文件不是 Excel 檔案。", MsgBoxStyle.Critical, "錯誤")
                    Case Language.en_US
                        MsgBox("The chosen file is not Excel document.", MsgBoxStyle.Critical, "Error")
                    Case Language.ja_JP
                        MsgBox("選択されたファイルは Excel ドキュメントではありません。", MsgBoxStyle.Critical, "エラー")
                End Select
            End If
            OpenFileDialog1.FileName = ""
        End If
    End Sub

    '浏览文件 2
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            If InStr(OpenFileDialog1.FileName.Length - 4, OpenFileDialog1.FileName, ".xls", CompareMethod.Text) <> 0 Then
                file2 = OpenFileDialog1.FileName
                TextBox2.Clear()
                TextBox2.AppendText(file2)
            Else
                Select Case lang
                    Case Language.zh_CN
                        MsgBox("所选文件不是 Excel 文档。", MsgBoxStyle.Critical, "错误")
                    Case Language.zh_TW
                        MsgBox("所選文件不是 Excel 檔案。", MsgBoxStyle.Critical, "錯誤")
                    Case Language.en_US
                        MsgBox("The chosen file is not Excel document.", MsgBoxStyle.Critical, "Error")
                    Case Language.ja_JP
                        MsgBox("選択されたファイルは Excel ドキュメントではありません。", MsgBoxStyle.Critical, "エラー")
                End Select
            End If
            OpenFileDialog1.FileName = ""
        End If
    End Sub

    '打开或重新打开两个文件
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If IsNothing(file1) OrElse IsNothing(file2) Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("请选择两个 Excel 文档。", MsgBoxStyle.Exclamation, "错误")
                Case Language.zh_TW
                    MsgBox("請選擇兩個 Excel 檔案。", MsgBoxStyle.Exclamation, "錯誤")
                Case Language.en_US
                    MsgBox("Please select two Excel documents.", MsgBoxStyle.Exclamation, "Error")
                Case Language.ja_JP
                    MsgBox("Excel ドキュメントを二つ選んでください。", MsgBoxStyle.Exclamation, "エラー")
            End Select
        Else
            If isOpened AndAlso isNeedSave Then
                Select Case lang
                    Case Language.zh_CN
                        If MsgBox("已修改的文档未保存，重新打开文档将导致数据丢失" & Chr(13) & "是否重新打开文档？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "打开文件").Equals(vbYes) Then
                            isNeedSave = False
                            OpenFilesAsTemp()
                        End If
                    Case Language.zh_TW
                        If MsgBox("已修改的檔案未保存，重新打開檔案將導致數據丟失" & Chr(13) & "是否重新打開檔案？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "打開文件").Equals(vbYes) Then
                            isNeedSave = False
                            OpenFilesAsTemp()
                        End If
                    Case Language.en_US
                        If MsgBox("The modified documents have not been saved. Reopening will cause data loss." & Chr(13) & "Do you want to reopen the documents?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "Open File").Equals(vbYes) Then
                            isNeedSave = False
                            OpenFilesAsTemp()
                        End If
                    Case Language.ja_JP
                        If MsgBox("変更されたドキュメントは保存されていません。ドキュメントを再度開くと、データが失われます。" & Chr(13) & "ドキュメントを再度開きますか。", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "ファイルを開く").Equals(vbYes) Then
                            isNeedSave = False
                            OpenFilesAsTemp()
                        End If
                End Select
            Else
                OpenFilesAsTemp()
            End If
        End If
    End Sub

    '表格 1 切换工作表
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '改变当前工作表
        sheet1 = book1.Worksheets(ComboBox1.SelectedItem)
        activeSheet1 = ComboBox1.SelectedIndex
    End Sub

    '表格 2 切换工作表
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '改变当前工作表
        sheet2 = book2.Worksheets(ComboBox2.SelectedItem)
        activeSheet2 = ComboBox2.SelectedIndex
    End Sub

    '帮助
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Select Case lang
            Case Language.zh_CN
                MsgBox("此工具对两个 Excel 文档的指定列中所有数据进行对比。相匹配的数据在两个文档中都用绿色背景色标注，无匹配的数据用红色背景色标注，单个文档内如有重复数据则用黄色字体标注。", 0, "说明")
            Case Language.zh_TW
                MsgBox("此工具對兩個 Excel 檔案的指定列中所有數據進行對比。相匹配的數據在兩個檔案中都用綠色背景色標注，無匹配的數據用紅色背景色標注，單個檔案內如有重複數據則用黃色字體標註。", 0, "說明")
            Case Language.en_US
                MsgBox("This tool compares all data in specified columns between two Excel documents. Matched data will be marked with green background color in both documents. Unmatched data will be marked with red background color. Repeated data within a document will be marked with yellow font color.", 0, "Description")
            Case Language.ja_JP
                MsgBox("このツールは両方の Excel ドキュメントの指定された列の全てのデータを比較します。一致するデータは両方のドキュメントに緑の背景色でマークされます。一致しないデータは赤の背景色でマークされます。単一のドキュメント内の重複したデータは黄色のフォントでマークされます。", 0, "説明")
        End Select
    End Sub

    '开始数据对比
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim runThread As Threading.Thread = New Threading.Thread(AddressOf RunCompare)
        runThread.Start()
    End Sub

    '另存为新文件
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim newFile1 As String, newFile2 As String
        DisableControls()


        If file1.Equals(file2) Then
            newFile1 = Replace(file1, ".xls", "_marked(1).xls", , , CompareMethod.Text)
            newFile2 = Replace(file2, ".xls", "_marked(2).xls", , , CompareMethod.Text)
        Else
            newFile1 = Replace(file1, ".xls", "_marked.xls", , , CompareMethod.Text)
            newFile2 = Replace(file2, ".xls", "_marked.xls", , , CompareMethod.Text)
        End If

        Dim temp As String = Button5.Text
        Select Case lang
            Case Language.zh_CN
                Button5.Text = "正在保存..."
            Case Language.zh_TW
                Button5.Text = "正在保存..."
            Case Language.en_US
                Button5.Text = "Saving..."
            Case Language.ja_JP
                Button5.Text = "保存中..."
        End Select

        '保存临时副本
        book1.Save()
        book2.Save()

        Try
            '复制临时副本到当前目录
            System.IO.File.Copy(tempFile1, newFile1, True)
            System.IO.File.Copy(tempFile2, newFile2, True)
        Catch ioe As System.IO.IOException
            Select Case lang
                Case Language.zh_CN
                    MsgBox("无法保存文件，目标文件已存在并且正在被使用。" & Chr(13) & "请关闭所有文件名中含有 ""_marked"" 的 Excel 文档后再保存。", MsgBoxStyle.Critical, "错误")
                Case Language.zh_TW
                    MsgBox("無法保存文件，目標文件已存在並且正在被使用。" & Chr(13) & "請關閉所有文件名中含有 ""_marked"" 的 Excel 檔案後再保存。", MsgBoxStyle.Critical, "錯誤")
                Case Language.en_US
                    MsgBox("Unable to save the files. The target files already exist and are currently being used." & Chr(13) & "Please close all Excel documents with ""_marked"" in the filename and then save again.", MsgBoxStyle.Critical, "Error")
                Case Language.ja_JP
                    MsgBox("ファイルを保存できません。ターゲットファイルが既に存在し、使用されています。" & Chr(13) & "ファイル名で ""_marked"" を含んでいる Excel ドキュメントを閉じてからもう一度保存してください。", MsgBoxStyle.Critical, "エラー")
            End Select
            Button5.Text = temp
            EnableControls()
            Exit Sub
        End Try

        Select Case lang
            Case Language.zh_CN
                MsgBox("新文件已另存至当前目录。")
            Case Language.zh_TW
                MsgBox("新文件已另存至當前目錄。")
            Case Language.en_US
                MsgBox("New files have been saved to current directory.")
            Case Language.ja_JP
                MsgBox("新しいファイルは現在のディレクトリに保存されました。")
        End Select

        isNeedSave = False
        EnableControls()
        Button5.Text = temp
    End Sub

    '退出
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Not isNeedSave Then
            Me.Dispose()
        Else
            Select Case lang
                Case Language.zh_CN
                    If MsgBox("已修改的文档未保存，直接退出将导致数据丢失。" & Chr(13) & "是否放弃保存并退出程序？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "关闭程序").Equals(vbYes) Then
                        Me.Dispose()
                    End If
                Case Language.zh_TW
                    If MsgBox("已修改的檔案未保存，直接退出將導致數據丟失。" & Chr(13) & "是否放棄保存並退出程序？", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "關閉程序").Equals(vbYes) Then
                        Me.Dispose()
                    End If
                Case Language.en_US
                    If MsgBox("The modified documents have not been saved. Reopening will cause data loss." & Chr(13) & "Do you want to reopen the documents?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "Open File").Equals(vbYes) Then
                        Me.Dispose()
                    End If
                Case Language.ja_JP
                    If MsgBox("変更されたドキュメントは保存されていません。保存せずに終了すると、データが失われます。" & Chr(13) & "保存せずにプログラムを終了しますか。", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2, "プログラムの終了").Equals(vbYes) Then
                        Me.Dispose()
                    End If
            End Select
        End If
    End Sub

    Private Sub Button4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.MouseEnter
        Select Case lang
            Case Language.zh_CN
                Button4.Text = "<<<<<<   开始数据对比   >>>>>>"
            Case Language.zh_TW
                Button4.Text = "<<<<<<   開始數據對比   >>>>>>"
            Case Language.en_US
                Button4.Text = "<<<<<<   Start Data Comparison   >>>>>>"
            Case Language.ja_JP
                Button4.Text = "<<<<<<   データの比較を始める   >>>>>>"
        End Select
    End Sub

    Private Sub Button4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.MouseLeave
        Select Case lang
            Case Language.zh_CN
                Button4.Text = "<<<   开始数据对比   >>>"
            Case Language.zh_TW
                Button4.Text = "<<<   開始數據對比   >>>"
            Case Language.en_US
                Button4.Text = "<<<   Start Data Comparison   >>>"
            Case Language.ja_JP
                Button4.Text = "<<<   データの比較を始める   >>>"
        End Select
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress, TextBox4.KeyPress
        '只能输入字母
        If Char.IsLetter(e.KeyChar) OrElse e.KeyChar = Chr(8) Then  'Chr(8) = 退格键
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress, TextBox6.KeyPress, TextBox7.KeyPress, TextBox8.KeyPress
        '只能输入数字
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Leave
        If TextBox3.Text.Length.Equals(2) AndAlso StrComp(TextBox3.Text, "IV").Equals(1) Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("列数超出范围 (A - IV)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("列數超出範圍 (A - IV)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Column number out of range (A - IV).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("列の数が範囲外になりました (A - IV)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox3.Focus()
            TextBox3.SelectAll()
        End If
    End Sub

    Private Sub TextBox3_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox3.MouseClick
        TextBox3.SelectAll()
    End Sub

    Private Sub TextBox4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Leave
        If TextBox4.Text.Length.Equals(2) AndAlso StrComp(TextBox4.Text, "IV").Equals(1) Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("列数超出范围 (A - IV)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("列數超出範圍 (A - IV)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Column number out of range (A - IV).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("列番号が範囲外になりました (A - IV)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox4.Focus()
            TextBox4.SelectAll()
        End If
    End Sub

    Private Sub TextBox4_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox4.MouseClick
        TextBox4.SelectAll()
    End Sub

    Private Sub TextBox5_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.Leave
        If Integer.Parse(TextBox5.Text) > 65536 OrElse Integer.Parse(TextBox5.Text) < 1 Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("行数超出范围 (1 - 65536)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("行數超出範圍 (1 - 65536)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Row number out of range (1 - 65536).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("行番号が範囲外になりました (1 - 65536)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox5.Focus()
            TextBox5.SelectAll()
        End If
    End Sub

    Private Sub TextBox5_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox5.MouseClick
        TextBox5.SelectAll()
    End Sub

    Private Sub TextBox6_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.Leave
        If Integer.Parse(TextBox6.Text) > 65536 OrElse Integer.Parse(TextBox6.Text) < 1 Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("行数超出范围 (1 - 65536)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("行數超出範圍 (1 - 65536)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Row number out of range (1 - 65536).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("行番号が範囲外になりました (1 - 65536)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox6.Focus()
            TextBox6.SelectAll()
        End If
    End Sub

    Private Sub TextBox6_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox6.MouseClick
        TextBox6.SelectAll()
    End Sub

    Private Sub TextBox7_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.Leave
        If Integer.Parse(TextBox7.Text) > 65536 OrElse Integer.Parse(TextBox7.Text) < 1 Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("行数超出范围 (1 - 65536)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("行數超出範圍 (1 - 65536)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Row number out of range (1 - 65536).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("行番号が範囲外になりました (1 - 65536)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox7.Focus()
            TextBox7.SelectAll()
        End If
    End Sub

    Private Sub TextBox7_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox7.MouseClick
        TextBox7.SelectAll()
    End Sub

    Private Sub TextBox8_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.Leave
        If Integer.Parse(TextBox8.Text) > 65536 OrElse Integer.Parse(TextBox8.Text) < 1 Then
            Select Case lang
                Case Language.zh_CN
                    MsgBox("行数超出范围 (1 - 65536)。", MsgBoxStyle.Exclamation, "数据错误")
                Case Language.zh_TW
                    MsgBox("行數超出範圍 (1 - 65536)。", MsgBoxStyle.Exclamation, "數據錯誤")
                Case Language.en_US
                    MsgBox("Row number out of range (1 - 65536).", MsgBoxStyle.Exclamation, "Data Error")
                Case Language.ja_JP
                    MsgBox("行番号が範囲外になりました (1 - 65536)。", MsgBoxStyle.Exclamation, "データエラー")
            End Select
            TextBox8.Focus()
            TextBox8.SelectAll()
        End If
    End Sub

    Private Sub TextBox8_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox8.MouseClick
        TextBox8.SelectAll()
    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Me.Opacity = TrackBar1.Value / 100
    End Sub

    '简体中文
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            lang = Language.zh_CN
            ChangeTextLanguage(lang)
        End If
    End Sub

    '繁体中文
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            lang = Language.zh_TW
            ChangeTextLanguage(lang)
        End If
    End Sub

    '英语
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            lang = Language.en_US
            ChangeTextLanguage(lang)
        End If
    End Sub

    '日语
    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            lang = Language.ja_JP
            ChangeTextLanguage(lang)
        End If
    End Sub

    '解除关键控件锁定
    Private Sub EnableControls()
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        Me.UseWaitCursor = False
    End Sub

    '锁定关键控件
    Private Sub DisableControls()
        Me.UseWaitCursor = True
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
    End Sub

    '显示进度条
    Private Sub ShowProgressBar()
        ProgressBar1.Visible = True
        Label14.Visible = False
        RadioButton1.Visible = False
        RadioButton2.Visible = False
        RadioButton3.Visible = False
        RadioButton4.Visible = False
    End Sub

    '隐藏进度条
    Private Sub HideProgressBar()
        ProgressBar1.Visible = False
        Label14.Visible = True
        RadioButton1.Visible = True
        RadioButton2.Visible = True
        RadioButton3.Visible = True
        RadioButton4.Visible = True
    End Sub

    '打开两个 Excel 文件（作为临时副本）
    Private Sub OpenFilesAsTemp()
        '关闭已打开的临时副本
        If Not IsNothing(book1) Then
            book1.Close()
            book1 = Nothing
        End If
        If Not IsNothing(book2) Then
            book2.Close()
            book2 = Nothing
        End If

        Try
            '生成原文件的临时副本（不占用原文件）
            System.IO.File.Copy(file1, tempFile1, True)
            System.IO.File.Copy(file2, tempFile2, True)
        Catch ioe As System.IO.IOException
            Select Case lang
                Case Language.zh_CN
                    MsgBox("无法创建临时副本，软件之前可能非正常退出。" & Chr(13) & "软件将自动重启，请确保目前所有打开的 Excel 文档都已保存。", MsgBoxStyle.Critical, "错误")
                Case Language.zh_TW
                    MsgBox("無法創建臨時副本，軟體之前可能非正常退出。" & Chr(13) & "軟體將自動重啟，請確保目前所有打開的 Excel 文檔都已保存。", MsgBoxStyle.Critical, "錯誤")
                Case Language.en_US
                    MsgBox("Unable to create temporary copies of the files. This may be caused by abnormal exiting of the software before." & Chr(13) & "The software will automatically restart. Please make sure that all currently opened Excel documents are saved.", MsgBoxStyle.Critical, "Error")
                Case Language.ja_JP
                    MsgBox("一時ファイルを作成できません。ソフトは前に非正常に終了した可能性があります。" & Chr(13) & "ソフトは後で自動的に再起動します。現在の全ての開いている Excel ドキュメントが保存されていることを確認してください。", MsgBoxStyle.Critical, "エラー")
            End Select
            Shell("taskkill /f /im excel.exe")
            Application.Restart()
        End Try

        '打开文件
        book1 = excel.Workbooks.Open(tempFile1)
        book2 = excel.Workbooks.Open(tempFile2)

        '获取所有工作表
        ComboBox1.Items.Clear()
        For i = 1 To book1.Worksheets.Count
            ComboBox1.Items.Add(book1.Worksheets(i).Name)
        Next i
        ComboBox2.Items.Clear()
        For j = 1 To book2.Worksheets.Count
            ComboBox2.Items.Add(book2.Worksheets(j).Name)
        Next j

        '下拉栏显示第一个工作表
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        GroupBox1.Text = file1.Remove(0, file1.LastIndexOf("\") + 1)
        GroupBox2.Text = file2.Remove(0, file1.LastIndexOf("\") + 1)

        EnableControls()
        isOpened = True
    End Sub

    '更改界面文字语言
    Private Sub ChangeTextLanguage(ByVal l As Language)
        Select Case l
            Case Language.zh_CN
                Me.Text = "Excel 数据对比工具"
                Label1.Text = "文件 1："
                Label2.Text = "文件 2："
                Label3.Text = "工作表："
                Label4.Text = "工作表："
                Label5.Text = "列："
                Label7.Text = "列："
                Label8.Text = "范围："
                Label9.Text = "范围："
                Label10.Text = "行"
                Label11.Text = "行"
                Label14.Text = "语言："
                Label15.Text = "透明度："
                Button1.Text = "浏览..."
                Button2.Text = "浏览..."
                Button4.Text = "<<<   开始数据对比   >>>"
                Button5.Text = "另存为新文件"
                Button6.Text = "退出"
                Button7.Text = "打开"
                If GroupBox1.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox1.Text = " 表格 1"
                End If
                If GroupBox2.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox2.Text = " 表格 2"
                End If
            Case Language.zh_TW
                Me.Text = "Excel 數據對比工具"
                Label1.Text = "文件 1："
                Label2.Text = "文件 2："
                Label3.Text = "工作表："
                Label4.Text = "工作表："
                Label5.Text = "列："
                Label7.Text = "列："
                Label8.Text = "範圍："
                Label9.Text = "範圍："
                Label10.Text = "行"
                Label11.Text = "行"
                Label14.Text = "語言："
                Label15.Text = "透明度："
                Button1.Text = "浏覽..."
                Button2.Text = "浏覽..."
                Button4.Text = "<<<   開始數據對比   >>>"
                Button5.Text = "另存爲新文件"
                Button6.Text = "退出"
                Button7.Text = "打開"
                If GroupBox1.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox1.Text = " 表格 1"
                End If
                If GroupBox2.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox2.Text = " 表格 2"
                End If
            Case Language.en_US
                Me.Text = "Excel Data Comparison Tool"
                Label1.Text = "File 1:"
                Label2.Text = "File 2:"
                Label3.Text = "Sheet:"
                Label4.Text = "Sheet:"
                Label5.Text = "Col:"
                Label7.Text = "Col:"
                Label8.Text = "Lines:"
                Label9.Text = "Lines:"
                Label10.Text = " "
                Label11.Text = " "
                Label14.Text = "Language:"
                Label15.Text = "Opacity:"
                Button1.Text = "Browse"
                Button2.Text = "Browse"
                Button4.Text = "<<<   Start Data Comparison   >>>"
                Button5.Text = "Save As New File"
                Button6.Text = "Exit"
                Button7.Text = "Open"
                If GroupBox1.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox1.Text = " Table 1"
                End If
                If GroupBox2.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox2.Text = " Table 2"
                End If
            Case Language.ja_JP
                Me.Text = "Excel データ比較ツール"
                Label1.Text = "ﾌｧｲﾙ 1："
                Label2.Text = "ﾌｧｲﾙ 2："
                Label3.Text = "ﾜｰｸｼｰﾄ："
                Label4.Text = "ﾜｰｸｼｰﾄ："
                Label5.Text = "列："
                Label7.Text = "列："
                Label8.Text = "範囲："
                Label9.Text = "範囲："
                Label10.Text = "行"
                Label11.Text = "行"
                Label14.Text = "言語："
                Label15.Text = "透明度："
                Button1.Text = "閲覧..."
                Button2.Text = "閲覧..."
                Button4.Text = "<<<   データの比較を始める   >>>"
                Button5.Text = "新しいファイルを保存"
                Button6.Text = "閉じる"
                Button7.Text = "開く"
                If GroupBox1.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox1.Text = " テーブル 1"
                End If
                If GroupBox2.Text.Chars(0).Equals(Chr(32)) Then
                    GroupBox2.Text = " テーブル 2"
                End If
        End Select
    End Sub

    '数据对比核心代码（线程调用）
    Private Sub RunCompare()
        DisableControls()
        Dim temp As String = Button4.Text
        Select Case lang
            Case Language.zh_CN
                Button4.Text = "······   正在运行   ······"
            Case Language.zh_TW
                Button4.Text = "······   正在運行   ······"
            Case Language.en_US
                Button4.Text = "······   Running   ······"
            Case Language.ja_JP
                Button4.Text = "······   実行中   ······"
        End Select

        Try
            If String.IsNullOrEmpty(TextBox3.Text) OrElse String.IsNullOrEmpty(TextBox4.Text) OrElse String.IsNullOrEmpty(TextBox5.Text) OrElse String.IsNullOrEmpty(TextBox6.Text) OrElse String.IsNullOrEmpty(TextBox7.Text) OrElse String.IsNullOrEmpty(TextBox8.Text) Then
                Select Case lang
                    Case Language.zh_CN
                        MsgBox("一个或多个参数未填写。", MsgBoxStyle.Exclamation, "错误")
                    Case Language.zh_TW
                        MsgBox("一個或多個參數未填寫。", MsgBoxStyle.Exclamation, "錯誤")
                    Case Language.en_US
                        MsgBox("One or more parameters are not filled.", MsgBoxStyle.Exclamation, "Error")
                    Case Language.ja_JP
                        MsgBox("一つ又は複数のパラメータは書き込まれていません。", MsgBoxStyle.Exclamation, "エラー")
                End Select
            ElseIf (Integer.Parse(TextBox5.Text) > Integer.Parse(TextBox6.Text)) OrElse (Integer.Parse(TextBox7.Text) > Integer.Parse(TextBox8.Text)) Then
                Select Case lang
                    Case Language.zh_CN
                        MsgBox("起始行数大于终止行数。", MsgBoxStyle.Exclamation, "数据错误")
                    Case Language.zh_TW
                        MsgBox("起始行數大于終止行數。", MsgBoxStyle.Exclamation, "數據錯誤")
                    Case Language.en_US
                        MsgBox("Starting line number is greater than ending line number.", MsgBoxStyle.Exclamation, "Data Error")
                    Case Language.ja_JP
                        MsgBox("開始の行番号は終了の行番号より大きくなっています。", MsgBoxStyle.Exclamation, "データエラー")
                End Select
            Else
                Dim match As Integer = 0  '匹配数据数量
                Dim repeat As Integer = 0  '重复数据数量
                Dim timer = New Stopwatch
                Dim array1() As CompareArray, array2() As CompareArray

                ProgressBar1.Value = ProgressBar1.Minimum
                ProgressBar1.Step = 1000 / 7
                ShowProgressBar()
                timer.Start()

                '保存 TextBox 中数据
                Dim column1 As String = TextBox3.Text
                Dim column2 As String = TextBox4.Text
                Dim rowBegin1 As Integer = Integer.Parse(TextBox5.Text)
                Dim rowEnd1 As Integer = Integer.Parse(TextBox6.Text)
                Dim rowBegin2 As Integer = Integer.Parse(TextBox7.Text)
                Dim rowEnd2 As Integer = Integer.Parse(TextBox8.Text)

                '导入单元格数据到数组
                ReDim array1(rowEnd1 - rowBegin1)
                ReDim array2(rowEnd2 - rowBegin2)
                For i = rowBegin1 To rowEnd1
                    With array1(i - rowBegin1)
                        .Data = Trim(sheet1.Range(column1 & i).Value)
                        .IsMatched = False
                        .IsRepeated = False
                    End With
                Next
                ProgressBar1.PerformStep()
                For j = rowBegin2 To rowEnd2
                    With array2(j - rowBegin2)
                        .Data = Trim(sheet2.Range(column2 & j).Value)
                        .IsMatched = False
                        .IsRepeated = False
                    End With
                Next
                ProgressBar1.PerformStep()

                '互相对比数据
                For i = 0 To array1.Length() - 1
                    If Not String.IsNullOrEmpty(array1(i).Data) Then  '空单元格不检测
                        For j = 0 To array2.Length() - 1
                            If array1(i).Data.Equals(array2(j).Data) Then
                                match = match + 1
                                array1(i).IsMatched = True
                                array2(j).IsMatched = True
                            End If
                        Next j
                    End If
                Next i
                ProgressBar1.PerformStep()

                '查找重复数据
                For i = 0 To array1.Length() - 2
                    If Not String.IsNullOrEmpty(array1(i).Data) Then
                        For k = i + 1 To array1.Length() - 1
                            If array1(i).Data.Equals(array1(k).Data) Then
                                repeat = repeat + 2
                                array1(i).IsRepeated = True
                                array1(k).IsRepeated = True
                            End If
                        Next
                    End If
                Next
                ProgressBar1.PerformStep()
                For j = 0 To array2.Length() - 2
                    If Not String.IsNullOrEmpty(array2(j).Data) Then
                        For k = j + 1 To array2.Length() - 1
                            If array2(j).Data.Equals(array2(k).Data) Then
                                repeat = repeat + 2
                                array2(j).IsRepeated = True
                                array2(k).IsRepeated = True
                            End If
                        Next
                    End If
                Next
                ProgressBar1.PerformStep()

                '标注单元格背景色与字体颜色
                For i = rowBegin1 To rowEnd1
                    If Not String.IsNullOrEmpty(array1(i - rowBegin1).Data) Then
                        If array1(i - rowBegin1).IsMatched Then
                            sheet1.Range(column1 & i).Interior.ColorIndex = 4  '绿色
                        Else
                            sheet1.Range(column1 & i).Interior.ColorIndex = 3  '红色
                        End If
                        If array1(i - rowBegin1).IsRepeated Then
                            sheet1.Range(column1 & i).Font.ColorIndex = 6  '黄色
                        End If
                    End If
                Next
                ProgressBar1.PerformStep()
                For j = rowBegin2 To rowEnd2
                    If Not String.IsNullOrEmpty(array2(j - rowBegin2).Data) Then
                        If array2(j - rowBegin2).IsMatched Then
                            sheet2.Range(column2 & j).Interior.ColorIndex = 4
                        Else
                            sheet2.Range(column2 & j).Interior.ColorIndex = 3
                        End If
                        If array2(j - rowBegin2).IsRepeated Then
                            sheet2.Range(column2 & j).Font.ColorIndex = 6
                        End If
                    End If
                Next
                ProgressBar1.PerformStep()

                isNeedSave = True
                timer.Stop()
                book1.Save()
                book2.Save()
                ProgressBar1.Value = ProgressBar1.Maximum

                Select Case lang
                    Case Language.zh_CN
                        MsgBox("对比完成！" & Chr(13) & "用时：" & timer.ElapsedMilliseconds / 1000 & " 秒" & Chr(13) & Chr(13) & "查找到 " & match & " 对匹配项。" & Chr(13) & "查找到 " & repeat & " 个重复项。", 0, "结果")
                    Case Language.zh_TW
                        MsgBox("對比完成！" & Chr(13) & "用時：" & timer.ElapsedMilliseconds / 1000 & " 秒" & Chr(13) & Chr(13) & "查找到 " & match & " 對匹配項。" & Chr(13) & "查找到 " & repeat & " 個重複項。", 0, "結果")
                    Case Language.en_US
                        MsgBox("Comparison finished!" & Chr(13) & "Time: " & timer.ElapsedMilliseconds / 1000 & " second(s)" & Chr(13) & Chr(13) & "Found " & match & " pair(s) of matched items." & Chr(13) & "Found " & repeat & " repeated item(s).", 0, "Results")
                    Case Language.ja_JP
                        MsgBox("比較完了！" & Chr(13) & "時間：" & timer.ElapsedMilliseconds / 1000 & " 秒" & Chr(13) & Chr(13) & "一致するアイテムは " & match & " 組あります。" & Chr(13) & "重複するアイテムは " & repeat & " 個あります。", 0, "结果")
                End Select
            End If

            HideProgressBar()
            EnableControls()
            Button4.Text = temp
        Catch comEx As System.Runtime.InteropServices.COMException
            Select Case lang
                Case Language.zh_CN
                    MsgBox("进程 EXCEL.EXE 被异常关闭，应用程序将重启。", MsgBoxStyle.Critical, "严重错误")
                Case Language.zh_TW
                    MsgBox("進程 EXCEL.EXE 被異常關閉，應用程序將重啓。", MsgBoxStyle.Critical, "嚴重錯誤")
                Case Language.en_US
                    MsgBox("Process EXCEL.EXE is terminated abnormally. The application will restart.", MsgBoxStyle.Critical, "Fatal Error")
                Case Language.ja_JP
                    MsgBox("プロセス EXCEL.EXE は異常に閉じられました。アプリケーションは再起動します。", MsgBoxStyle.Critical, "致命的なエラー")
            End Select
            Shell("taskkill /f /im excel.exe")
            Application.Restart()
        Catch ex As Exception
            Select Case lang
                Case Language.zh_CN
                    MsgBox("出现未知错误，程序即将终止。", MsgBoxStyle.Critical, "严重错误")
                Case Language.zh_TW
                    MsgBox("出現未知錯誤，程序即將終止。", MsgBoxStyle.Critical, "嚴重錯誤")
                Case Language.en_US
                    MsgBox("An unknown error has occurred. The program will be terminated.", MsgBoxStyle.Critical, "Fatal Error")
                Case Language.ja_JP
                    MsgBox("不明なエラーが発生し、プログラムが終了します。", MsgBoxStyle.Critical, "致命的なエラー")
            End Select
            Shell("taskkill /f /im excel.exe")
            Me.Dispose()
        End Try
    End Sub
End Class