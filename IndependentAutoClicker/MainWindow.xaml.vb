Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports Gma.System.MouseKeyHook

Class MainWindow

    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202


    <DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="WindowFromPoint", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function WindowFromPoint(point As POINT) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ScreenToClient(ByVal hWnd As IntPtr, ByRef lpPoint As POINT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef processId As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Shared Function GetCursorPos(ByRef lpPoint As POINT) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure POINT
        Public X As Integer
        Public Y As Integer

        Public Sub New(ByVal X As Integer, ByVal Y As Integer)
            Me.X = X
            Me.Y = Y
        End Sub
    End Structure

    Private GlobalHook As IKeyboardMouseEvents

    Private Hwnd As IntPtr

    Private relLoc As POINT

    Private hotkey As Keys

    Private setHotkey As Boolean

    Private tokenSrc As CancellationTokenSource

    Private watch As Stopwatch = New Stopwatch
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        cmb_cps.ItemsSource = Enumerable.Range(1, 100)

        GlobalHook = Hook.GlobalEvents

        AddHandler GlobalHook.KeyDown, AddressOf hookKeyDown
        btn_clear.IsEnabled = False

    End Sub

    Private Sub clickWindow(clicks As Integer, token As CancellationToken)
        watch.Restart()
        Dim i As Long = 1
        Dim ms As Double = 1000
        While (Not token.IsCancellationRequested)
            If i Mod clicks = 0 Then
                ms = (ms + watch.ElapsedMilliseconds) / 2
                watch.Restart()
            End If
            SendMouseMsg(Hwnd, WM_LBUTTONDOWN, relLoc)
            SendMouseMsg(Hwnd, WM_LBUTTONUP, relLoc)
            Thread.Sleep(Math.Max(Math.Round((1000 - (ms - 1000)) / clicks), 0))
            'Await Task.Delay(Math.Max(Math.Round(1000 / clicks) - watch.ElapsedMilliseconds, 0))
            i += 1
        End While
        watch.Stop()
    End Sub

    Private Sub hookKeyDown(sender As Object, e As KeyEventArgs)
        If setHotkey Then
            If e.KeyCode = Keys.Escape Then
                hotkey = Keys.None
                lbl_htky.Content = "None"
            Else
                hotkey = e.KeyCode
                lbl_htky.Content = hotkey.ToString
            End If
            setHotkey = False
        ElseIf e.KeyCode = hotkey Then
            toggleClicker()
        End If
    End Sub

    Private Sub toggleClicker()
        If Not cmb_cps.SelectedItem Is Nothing Then
            If tokenSrc Is Nothing Then
                tokenSrc = New CancellationTokenSource
                Dim clicks As Integer = cmb_cps.SelectedItem
                Task.Run(Sub() clickWindow(clicks, tokenSrc.Token), tokenSrc.Token)
            Else
                tokenSrc.Cancel()
                tokenSrc = Nothing
            End If
        Else
            MsgBox("U have to choose a Cps u idiot!!")
        End If

    End Sub

    Private Sub btn_choose_Click(sender As Object, e As RoutedEventArgs) Handles btn_choose.Click

        AddHandler GlobalHook.MouseDownExt, AddressOf hookMouseDown

    End Sub

    Private Sub hookMouseDown(sender As Object, e As MouseEventExtArgs)
        e.Handled = True
        RemoveHandler GlobalHook.MouseDownExt, AddressOf hookMouseDown
        Dim pos As POINT
        GetCursorPos(pos)
        relLoc = pos
        Hwnd = WindowFromPoint(pos)
        ScreenToClient(Hwnd, relLoc)
        Dim pid As UInteger
        GetWindowThreadProcessId(Hwnd, pid)
        targetLbl.Content = $"{Process.GetProcessById(pid).ProcessName} at {relLoc.X}x{relLoc.Y}"
        btn_clear.IsEnabled = True
    End Sub

    Private Sub SendMouseMsg(handle As IntPtr, type As Integer, pos As POINT)


        SendMessage(handle, type, New IntPtr(1), New IntPtr(((pos.Y << 16) Or (pos.X And &HFFFF))))


    End Sub

    Private Sub btn_setHtky_Click(sender As Object, e As RoutedEventArgs) Handles btn_setHtky.Click
        setHotkey = True
    End Sub

    Private Sub btn_clear_Click(sender As Object, e As RoutedEventArgs) Handles btn_clear.Click
        If tokenSrc IsNot Nothing Then
            tokenSrc.Cancel()
            tokenSrc = Nothing
        End If
        Hwnd = Nothing
        relLoc = New POINT
    End Sub
End Class
