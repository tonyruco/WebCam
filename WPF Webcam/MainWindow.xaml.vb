Imports Microsoft.Expression.Encoder.Devices
Imports WebcamControl
Imports System.IO
Imports Microsoft.Expression.Encoder.Live
Imports Microsoft.Expression.Encoder
Imports System.Drawing
Imports System.Drawing.Imaging

Class MainWindow

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Dim binding_1 As New Binding("SelectedValue")
        binding_1.Source = VideoDevicesComboBox
        WebcamCtrl.SetBinding(Webcam.VideoDeviceProperty, binding_1)

        Dim binding_2 As New Binding("SelectedValue")
        binding_2.Source = AudioDevicesComboBox
        WebcamCtrl.SetBinding(Webcam.AudioDeviceProperty, binding_2)

        ' Create directory for saving video files
        Dim videoPath As String = "C:\VideoClips"

        If Not Directory.Exists(videoPath) Then
            Directory.CreateDirectory(videoPath)
        End If
        ' Create directory for saving image files
        Dim imagePath As String = "C:\WebcamSnapshots"

        If Not Directory.Exists(imagePath) Then
            Directory.CreateDirectory(imagePath)
        End If

        ' Set some properties of the Webcam control
        WebcamCtrl.VideoDirectory = videoPath
        WebcamCtrl.ImageDirectory = imagePath
        WebcamCtrl.FrameRate = 30
        WebcamCtrl.FrameSize = New Size(640, 480)

        ' Find available a/v devices
        Dim videoDevices = EncoderDevices.FindDevices(EncoderDeviceType.Video)
        Dim audioDevices = EncoderDevices.FindDevices(EncoderDeviceType.Audio)
        VideoDevicesComboBox.ItemsSource = videoDevices
        AudioDevicesComboBox.ItemsSource = audioDevices
        VideoDevicesComboBox.SelectedIndex = 0
        AudioDevicesComboBox.SelectedIndex = 0
    End Sub

    Private Sub StartCaptureButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Display webcam video
        Try
            WebcamCtrl.StartPreview()
        Catch ex As Microsoft.Expression.Encoder.SystemErrorException
            MessageBox.Show("Device is in use by another application")
        End Try
    End Sub

    Private Sub StopCaptureButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Stop the display of webcam video
        WebcamCtrl.StopPreview()
    End Sub

    Private Sub StartRecordingButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Start recording of webcam video
        WebcamCtrl.StartRecording()
    End Sub

    Private Sub StopRecordingButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Stop recording of webcam video
        WebcamCtrl.StopRecording()
    End Sub

    Private Sub TakeSnapshotButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Take snapshot of webcam video
        WebcamCtrl.TakeSnapshot()
    End Sub
End Class
