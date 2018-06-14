Public Class Form1
    'マウスのクリック位置を記憶
    Private mousePoint As Point

    'ページスイッチ(0:トップ,1:シャットダウン,2:アラーム,3...)
    Private pSW As Integer

    'カルチャを en-US にしてDataTimeを文字列に変換する
    Dim ci As New System.Globalization.CultureInfo("en-US")

    'シャットダウンする時間
    Private ShutdownDT As DateTime

    'タイマーが鳴る時間
    Private TimerDT As DateTime

    '時間と時刻のSW(0:時間,1:時刻)
    Private TimeSW As Integer

    'シャットダウン関係 --------------------------------------------------------------
    Public Enum ExitWindows
        EWX_LOGOFF = &H0
        EWX_SHUTDOWN = &H1
        EWX_REBOOT = &H2
        EWX_POWEROFF = &H8
        EWX_RESTARTAPPS = &H40
        EWX_FORCE = &H4
        EWX_FORCEIFHUNG = &H10
    End Enum

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ExitWindowsEx(ByVal uFlags As ExitWindows,
    ByVal dwReason As Integer) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr,
    ByVal DesiredAccess As Integer,
    ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(ByVal hHandle As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True,
    CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function LookupPrivilegeValue(ByVal lpSystemName As String,
    ByVal lpName As String,
    ByRef lpLuid As Long) As Boolean
    End Function

    <System.Runtime.InteropServices.StructLayout(
    System.Runtime.InteropServices.LayoutKind.Sequential, Pack:=1)>
    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As Integer
        Public Luid As Long
        Public Attributes As Integer
    End Structure

    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr,
    ByVal DisableAllPrivileges As Boolean,
    ByRef NewState As TOKEN_PRIVILEGES,
    ByVal BufferLength As Integer,
    ByVal PreviousState As IntPtr,
    ByVal ReturnLength As IntPtr) As Boolean
    End Function

    'シャットダウンするためのセキュリティ特権を有効にする
    Public Shared Sub AdjustToken()
        Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
        Const TOKEN_QUERY As Integer = &H8
        Const SE_PRIVILEGE_ENABLED As Integer = &H2
        Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"

        If Environment.OSVersion.Platform <> PlatformID.Win32NT Then
            Return
        End If

        Dim procHandle As IntPtr = GetCurrentProcess()

        'トークンを取得する
        Dim tokenHandle As IntPtr
        OpenProcessToken(procHandle,
        TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, tokenHandle)
        'LUIDを取得する
        Dim tp As New TOKEN_PRIVILEGES()
        tp.Attributes = SE_PRIVILEGE_ENABLED
        tp.PrivilegeCount = 1
        LookupPrivilegeValue(Nothing, SE_SHUTDOWN_NAME, tp.Luid)
        '特権を有効にする
        AdjustTokenPrivileges(tokenHandle, False, tp, 0, IntPtr.Zero, IntPtr.Zero)

        '閉じる
        CloseHandle(tokenHandle)
    End Sub
    'シャットダウン関係終わり---------------------------------------------------------------

    'Enterでピープオンならないようにするためのもの
    <System.Security.Permissions.UIPermission(
    System.Security.Permissions.SecurityAction.Demand,
    Window:=System.Security.Permissions.UIPermissionWindow.AllWindows)>
    Protected Overrides Function ProcessDialogKey(
        ByVal keyData As Keys) As Boolean
        'TextBox1でEnterを押してもビープ音が鳴らないようにする
        If TextBox1.Focused AndAlso
        (keyData And Keys.KeyCode) = Keys.Enter Then
            Return True
        End If
        Return MyBase.ProcessDialogKey(keyData)
    End Function

    'マウスのボタンが押されたとき
    Private Sub Form1_MouseDown(ByVal sender As Object,
        ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles MyBase.MouseDown, PictureBox1.MouseDown, PictureBox2.MouseDown
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            '位置を記憶する
            mousePoint = New Point(e.X, e.Y)
        End If
    End Sub

    'マウスが動いたとき
    Private Sub Form1_MouseMove(ByVal sender As Object,
        ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles MyBase.MouseMove, PictureBox1.MouseMove, PictureBox2.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            Me.Left += e.X - mousePoint.X
            Me.Top += e.Y - mousePoint.Y
        End If
    End Sub

    'フォームロード
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'フォームの表示位置を設定する。
        Me.Left = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 640
        Me.Top = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - 340

        'ページスイッチをトップにする
        pSW = 0

        '時間、時刻スイッチを時間にする
        TimeSW = 0

        'ページが違うコントロールを非表示にしておく
        'タイマー系
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Button1.Visible = False
        Button2.Visible = False
        PictureBox5.Visible = False

        'シャットダウン用
        Label5.Visible = False

        'アラーム用
        Label6.Visible = False

        '検索用
        Label7.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        TextBox1.Visible = False

        'メモ用
        TextBox2.Visible = False


        'ピクチャボックス2の子にする
        'トップ表示系
        PictureBox2.Controls.Add(PictureBox3)
        PictureBox2.Controls.Add(PictureBox4)
        PictureBox2.Controls.Add(PictureBox6)
        PictureBox2.Controls.Add(PictureBox7)
        PictureBox2.Controls.Add(PictureBox8)
        PictureBox2.Controls.Add(Label1)
        PictureBox2.Controls.Add(Label2)

        'タイマー系
        PictureBox2.Controls.Add(ComboBox1)
        PictureBox2.Controls.Add(ComboBox2)
        PictureBox2.Controls.Add(ComboBox3)
        PictureBox2.Controls.Add(ComboBox4)
        PictureBox2.Controls.Add(Label3)
        PictureBox2.Controls.Add(Label4)
        PictureBox2.Controls.Add(Label5)
        PictureBox2.Controls.Add(Label6)
        PictureBox2.Controls.Add(Button1)
        PictureBox2.Controls.Add(Button2)
        PictureBox2.Controls.Add(PictureBox5)

        '検索系
        PictureBox2.Controls.Add(Label7)
        PictureBox2.Controls.Add(Button3)
        PictureBox2.Controls.Add(Button4)
        PictureBox2.Controls.Add(TextBox1)

        'メモ用
        PictureBox2.Controls.Add(TextBox2)

    End Sub

    'アプリケーションを終了する。
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        End
    End Sub

    'シャットダウンタイマーページ
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        'ページスイッチをシャットダウンページにする
        pSW = 1

        'トップページを非表示にする
        HideTopControls()
        'シャットダウンページを表示する。
        DispOtherPageControls()
    End Sub

    'アラームページ
    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        'ページスイッチをシャットダウンページにする
        pSW = 2

        'トップページを非表示にする
        HideTopControls()
        'アラームページを表示する。
        DispOtherPageControls()
    End Sub

    'Web検索ページ
    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        'ページスイッチをWeb検索ページにする
        pSW = 3

        'トップページを非表示にする
        HideTopControls()
        'Web検索ページを表示する。
        DispOtherPageControls()
    End Sub

    'メモページ
    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        'ページスイッチをメモページにする。
        pSW = 4

        'トップページを非表示にする
        HideTopControls()
        'Web検索ページを表示する。
        DispOtherPageControls()
    End Sub

    'トップページを非表示にする。
    Private Sub HideTopControls()
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False

        Label1.Visible = False
        Label2.Visible = False
    End Sub

    'トップページを表示する
    Private Sub DispTopControls()
        PictureBox3.Visible = True
        PictureBox4.Visible = True
        PictureBox6.Visible = True
        PictureBox7.Visible = True
        PictureBox8.Visible = True

        Label1.Visible = True
        Label2.Visible = True
    End Sub

    'トップページ以外を非表示にする。
    Private Sub HideOtherPageControls()

        If pSW = 1 Then
            ComboBox1.Visible = False
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Button1.Visible = False
            Button2.Visible = False
            PictureBox5.Visible = False
        ElseIf pSW = 2 Then
            ComboBox1.Visible = False
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            Label6.Visible = False
            Button1.Visible = False
            Button2.Visible = False
            PictureBox5.Visible = False
        ElseIf pSW = 3 Then
            Label7.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            TextBox1.Visible = False
        ElseIf pSW = 4 Then
            TextBox2.Visible = False
            Button4.Visible = False
        End If

    End Sub

    'トップぺージ以外を表示する。
    Private Sub DispOtherPageControls()
        If pSW = 1 Then
            ComboBox1.Visible = True
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            Button1.Visible = True
            Button2.Visible = True
            PictureBox5.Visible = True

            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 1
            ComboBox3.SelectedIndex = 0
            ComboBox4.SelectedIndex = 0

            If TimeSW = 0 Then
                ComboBox3.Visible = False
                ComboBox4.Visible = False
            ElseIf TimeSW = 1 Then
                ComboBox1.Visible = False
            End If

        ElseIf pSW = 2 Then
            ComboBox1.Visible = True
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            ComboBox4.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label6.Visible = True
            Button1.Visible = True
            Button2.Visible = True
            PictureBox5.Visible = True

            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 1
            ComboBox3.SelectedIndex = 0
            ComboBox4.SelectedIndex = 0

            If TimeSW = 0 Then
                ComboBox3.Visible = False
                ComboBox4.Visible = False
            ElseIf TimeSW = 1 Then
                ComboBox1.Visible = False
            End If

        ElseIf pSW = 3 Then
            Label7.Visible = True
            Button3.Visible = True
            Button4.Visible = True
            TextBox1.Visible = True

        ElseIf pSW = 4 Then
            TextBox2.Text = IO.File.ReadAllText("nicomemo.txt",
                                                System.Text.Encoding.GetEncoding("unicode"))
            TextBox2.Visible = True
            Button4.Visible = True
        End If

    End Sub

    'キャンセルボタン
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'オーディオリソースを取り出す
        Dim strm As System.IO.Stream = My.Resources.pipi
        '再生する
        My.Computer.Audio.Play(strm, AudioPlayMode.Background)

        'トップのコントロールを表示する。
        DispTopControls()

        'トップじゃないページのコントロールを非表示にする。
        HideOtherPageControls()

        '現在のページごとの処理
        If pSW = 1 Then
            'シャットダウンページのとき
            Label1.Text = "The shutdown timer is off"
            'シャットダウンチェックタイマを止める
            Timer1.Stop()
        ElseIf pSW = 2 Then
            'アラームページのとき
            Label2.Text = "The alarm clock is off"
            'シャットダウンチェックタイマを止める
            Timer2.Stop()
        End If


        'ページSWをトップページにする
        pSW = 0
    End Sub

    'OKボタン
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'バックカラーとトランスペアレンシーをOKにこちゃんに合わせる。
        Me.BackColor = Color.FromArgb(255, 200, 255)
        'バックカラーが透過色になるようにする。
        Me.TransparencyKey = Color.FromArgb(255, 200, 255)

        'にこちゃんの表示をOKにする。
        PictureBox1.Image = My.Resources.nicoOK

        'にこちゃんの表示を戻す用のタイマー
        Timer3.Start()

        'オーディオリソースを取り出す
        Dim strm As System.IO.Stream = My.Resources.piko
        '再生する
        My.Computer.Audio.Play(strm, AudioPlayMode.Background)

        'トップのコントロールを表示する。
        DispTopControls()

        'トップじゃないページのコントロールを非表示にする。
        HideOtherPageControls()

        '現在のページごとの処理
        If pSW = 1 Then
            'シャットダウンページのとき
            If TimeSW = 0 Then
                '時間のとき
                Dim ts1 As New TimeSpan(ComboBox1.SelectedItem, ComboBox2.SelectedItem, 0)
                ShutdownDT = System.DateTime.Now + ts1
            ElseIf TimeSW = 1 Then
                '時刻のとき
                ShutdownDT = New DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day,
                           ComboBox3.SelectedItem, ComboBox2.SelectedItem, 0)
                '午後だったら12時間足す。
                If ComboBox4.SelectedItem = "PM" Then
                    Dim ts3 As New TimeSpan(0, 12, 0, 0)
                    ShutdownDT = ShutdownDT + ts3
                End If
                '現在時刻より前になっていないか。
                If ShutdownDT < System.DateTime.Now Then
                    Dim ts2 As New TimeSpan(1, 0, 0, 0)
                    ShutdownDT = ShutdownDT + ts2
                End If
            End If
            Label1.Text = "I would shut down at " + ShutdownDT.ToString("t", ci)

            Timer1.Start()

        ElseIf pSW = 2 Then
            'アラームページのとき
            If TimeSW = 0 Then
                '時間のとき
                Dim ts1 As New TimeSpan(ComboBox1.SelectedItem, ComboBox2.SelectedItem, 0)
                TimerDT = System.DateTime.Now + ts1
            ElseIf TimeSW = 1 Then
                '時刻のとき
                TimerDT = New DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day,
                               ComboBox3.SelectedItem, ComboBox2.SelectedItem, 0)
                '午後だったら12時間足す。
                If ComboBox4.SelectedItem = "PM" Then
                    Dim ts3 As New TimeSpan(0, 12, 0, 0)
                    TimerDT = TimerDT + ts3
                End If
                '現在時刻より前になっていないか。
                If TimerDT < System.DateTime.Now Then
                    Dim ts2 As New TimeSpan(1, 0, 0, 0)
                    TimerDT = TimerDT + ts2
                End If
            End If
            Label2.Text = "I would ring the alarm at " + TimerDT.ToString("t", ci)

            Timer2.Start()
        End If

        'ページSWをトップページにする
        pSW = 0
    End Sub

    'シャットダウンの時間をチェックするタイマ
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ShutdownDT < System.DateTime.Now Then
            'シャットダウンする
            AdjustToken()
            ExitWindowsEx(ExitWindows.EWX_POWEROFF, 0)
            Timer1.Stop()
        End If
    End Sub

    'アラームの時間をチェックするタイマ
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If TimerDT < System.DateTime.Now Then
            'オーディオリソースを取り出す
            Dim strm As System.IO.Stream = My.Resources.alarm
            '再生する
            My.Computer.Audio.Play(strm, AudioPlayMode.WaitToComplete)

            Label2.Text = "The alarm clock is off"
            Timer2.Stop()
        End If
    End Sub

    '時間、時刻切り替え
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If TimeSW = 0 Then
            TimeSW = 1
            Label3.Text = ":"
            Label4.Text = ""
            Label5.Text = "I would shut down at"
            Label6.Text = "I would ring the alarm at"
            ComboBox2.Left = 125
            ComboBox1.Visible = False
            ComboBox3.Visible = True
            ComboBox4.Visible = True
        ElseIf TimeSW = 1 Then
            TimeSW = 0
            Label3.Text = "hours"
            Label4.Text = "min"
            Label5.Text = "I would shut down after"
            Label6.Text = "I would ring the alarm after"
            ComboBox2.Left = 165
            ComboBox1.Visible = True
            ComboBox3.Visible = False
            ComboBox4.Visible = False
        End If

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 1
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
    End Sub

    'サーチボタン
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'オーディオリソースを取り出す
        Dim strm As System.IO.Stream = My.Resources.piko
        '再生する
        My.Computer.Audio.Play(strm, AudioPlayMode.Background)

        System.Diagnostics.Process.Start("https://www.google.co.jp/search?&q=" & TextBox1.Text)
        TextBox1.Text = ""
    End Sub

    '戻るボタン
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'オーディオリソースを取り出す
        Dim strm As System.IO.Stream = My.Resources.pipi
        '再生する
        My.Computer.Audio.Play(strm, AudioPlayMode.Background)

        'トップのコントロールを表示する。
        DispTopControls()

        'トップじゃないページのコントロールを非表示にする。
        HideOtherPageControls()

        'メモページから戻るときにはメモをtxtに保存する。
        If pSW = 4 Then
            Dim sw As New System.IO.StreamWriter("nicomemo.txt",
    False,
    System.Text.Encoding.GetEncoding("utf-16"))
            'TextBox1.Textの内容を書き込む
            sw.Write(TextBox2.Text)
            '閉じる
            sw.Close()
        End If

        'ページSWをトップページにする
        pSW = 0
    End Sub

    'サーチページのテキストボックスでエンターキーが押されたら
    Private Sub TextBox1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox1.PreviewKeyDown
        If Keys.Enter = e.KeyCode Then
            'オーディオリソースを取り出す
            Dim strm As System.IO.Stream = My.Resources.piko
            '再生する
            My.Computer.Audio.Play(strm, AudioPlayMode.Background)

            System.Diagnostics.Process.Start("https://www.google.co.jp/search?&q=" & TextBox1.Text)
            TextBox1.Text = ""
        End If
    End Sub

    'OKのにこちゃん表示から通常のにこちゃんに戻す。
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Timer3.Stop()

        'バックカラーを変える。
        Me.BackColor = Color.FromArgb(255, 199, 254)
        'バックカラーが透過色になるようにする。
        Me.TransparencyKey = Color.FromArgb(255, 199, 254)

        PictureBox1.Image = My.Resources.nicoloop
    End Sub

    'にこにークリック
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'にこちゃんの表示をうなずかせる。
        PictureBox1.Image = My.Resources.nicoclick

        'にこちゃんの表示を戻す用のタイマー
        Timer3.Start()
    End Sub
End Class
