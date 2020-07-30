Imports System.IO
Imports System.IO.Ports
Imports GMap.NET
Imports GMap.NET.WindowsForms
Imports GMap.NET.AccessMode
Imports System.Threading
Imports vb = Microsoft.VisualBasic
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form1
    Dim myPort As Array 'inisialisasi untuk mendapatkan penamaan port
    Dim readBuffer As String
    Dim hitungmundur As Integer

    Dim baca As String 'variable untuk baca serial Port
    Dim baca_clear As String ' variable yg sudah di hapus karakter CR+LF
    Dim data_masuk() As String 'Array untuk membagi part data serial yang masuk

    'variable untuk tanggal dan jam
    Dim tanggal As String
    Dim Jam As String
    Dim jumlah_koma As Integer

    Delegate Sub SetTextCallBack(ByVal [text] As String)
    Private Delegate Sub UpdateFormDelegate() 'ini inisialisasi untuk memuat update terdelegate, atau mencegah aplikasi tidak terjadi lagging disebabkan oleh cepatnya data yang dikirim dari payload
    Private UpdateFormDelegate1 As UpdateFormDelegate
    Private data_pertama As Boolean
    'Private lat, lon As Double
    'Private WithEvents payload_marker As Markers.GMarkerCross
    'atas diganti ini
    Private WithEvents payload_marker As Markers.GMarkerGoogle
    '------------
    Private WithEvents payload_layer As GMapOverlay

    'tambahan
    Private WithEvents marker_layer As GMapOverlay 'marker_layer adalah tanda posisi GPS dalam bentuk titik titik
    Private WithEvents dgv1 As DataGridView
    '------------------

    'Label Untuk Parameter Atmosfer
    Dim ID As String
    Dim Wkt As String
    Dim Altitude As String
    Dim Temperature As String
    Dim Humidity As String
    Dim Pressure As String
    Dim WC As String
    Dim WS As String
    Dim arah As String
    Dim Latitude As String
    Dim Longitude As String
    Dim WC1 As Double
    Dim WS1 As Double

    'Label Untuk Masukan Antenna Tracker
    Dim IDgps As String
    Dim wkt2 As String
    Dim Latitude2 As String
    Dim HeadingLatitude2 As String
    Dim Longitude2 As String
    Dim HeadingLongitude2 As String
    Dim GPSstatus As String
    Dim Sats As String
    Dim HDOP As String
    Dim satuan1 As String
    Dim AltSL As String
    Dim satuan2 As String
    Dim kosong As String
    Dim kosong2 As String
    Dim CKS As String
    Dim IDgps2 As String
    Dim kosong3 As String
    Dim T As String
    Dim kosong4 As String
    Dim M As String
    Dim GSpeed As String
    Dim N As String
    Dim GSpeed2 As String
    Dim K As String
    Dim CKS2 As String

    'Tambahan
    Private rc6in As Double
    Private tth As Double
    '------------

    'Pertama kali Program di buka, dia akan load ini semua parameter
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Longitude = 0
        Latitude = 0

        'Tambahan
        rc6in = 1000
        tth = 1500 ' ganti sesuai data lapangan
        '-------------------

        myPort = IO.Ports.SerialPort.GetPortNames()
        PortComboBox1.Items.AddRange(myPort)
        data_pertama = True

        'tambahan:
        marker_layer = New GMapOverlay("circle_layer")
        GMapControl1.MinZoom = 5
        GMapControl1.MaxZoom = 20
        GMapControl1.Zoom = 17
        '----------------

        GMapControl1.Position = New PointLatLng(Latitude, Longitude)
        GMapControl1.MapProvider = MapProviders.BingHybridMapProvider.Instance
        'atas diganti ini
        'GMapControl1.MapProvider = MapProviders.BingSatelliteMapProvider.Instance
        GMapControl1.Manager.Mode = AccessMode.ServerAndCache
        payload_layer = New GMapOverlay("payload_layer")
        'payload_marker = New Markers.GMarkerCross(New PointLatLng(Latitude, Longitude), My.Resources.icon_payload)
        'atas diganti ini
        payload_marker = New Markers.GMarkerGoogle(New PointLatLng(Latitude, Longitude), My.Resources.icon_payload)
        '----------
        payload_layer.Markers.Add(payload_marker)
        GMapControl1.Overlays.Add(payload_layer)

        'Tambahan
        GMapControl1.Overlays.Add(marker_layer)
        '-------------

        GMapControl1.UpdateMarkerLocalPosition(payload_marker)
        GMapControl1.Invalidate()
        TableLayoutPanel2.Visible = True
        Button1.BackgroundImage = My.Resources.button_1

        Me.Chart2.Series("Series1").Points.AddXY(0, 0)
        Me.Chart3.Series("Series1").Points.AddXY(0, 0)
        Me.Chart4.Series("Series1").Points.AddXY(0, 0)
        Me.Chart5.Series("Series1").Points.AddXY(0, 0)
        Me.Chart6.Series("Series1").Points.AddXY(0, 0)

    End Sub

    'Pembacaan Status ADA atau TIDAKnya Komunikasi Serial yang sedang berlangsung'
    Private Sub SetupCOMPortButton3_Click(sender As Object, e As EventArgs) Handles SetupCOMPortButton3.Click
        If SerialPort1.IsOpen Then '<---- Jika Komunikasi Serial Port1 terbuka, maka menjalankan perintah dibawah'
            SerialPort1.Close()
            SetupCOMPortButton3.BackgroundImage = My.Resources.button_com1
            Label25.Text = "Disconnected"
            Label25.ForeColor = Color.Red
            Timer1.Enabled = False
            Label37.Text = "OFF"
            Label37.ForeColor = Color.Red
            Label33.Text = Date.Now.ToString("H:mm:ss")
        Else
            With SerialPort1 '<---- Jika terdapat Komunikasi Serial, maka...'
                SerialPort1.PortName = PortComboBox1.Text '<---SerialPort1 PortName membaca isi dari PortComboBox1 (COM XX)'
                SerialPort1.BaudRate = BaudComboBox2.Text '<---SerialPort1 BaudRate membaca isi dari PortComboBox2 (Nilai BaudRate)'
                'SerialPort1.DataBits = 8
                'SerialPort1.Parity = Parity.None
                'SerialPort1.StopBits = StopBits.One
                'SerialPort1.Handshake = Handshake.None
                'SerialPort1.Encoding = System.Text.Encoding.Default
                Label25.Text = "Connected"
                Label25.ForeColor = Color.LimeGreen
            End With
            Try
                SerialPort1.Open()
                SetupCOMPortButton3.BackgroundImage = My.Resources.button_com2
                Timer1.Enabled = True
                Label37.Text = "ON"
                Label37.ForeColor = Color.LimeGreen
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                Exit Sub 'exit sub if the connection fails
            End Try
            Threading.Thread.Sleep(200) 'wait 0.2 sec for port to open
        End If

    End Sub

    Private Sub SerialPort1_DataReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Invoke(New myDelegate(AddressOf updateTextBox), New Object() {})
        Invoke(New myDelegate(AddressOf updateTextBox3), New Object() {})
    End Sub

    Public Delegate Sub myDelegate()

    ' baca data serial yang masuk
    Public Sub updateTextBox()
        baca = SerialPort1.ReadLine()

        'data yang masuk diakhiri dengan CR + LF
        'replace atau ganti karakter \n dan \r dengan empty string
        baca_clear = Replace(Replace(baca, vbLf, ""), vbCr, "")

        'kemudian pisah tiap bagian dengan pemisah, tanda koma menjadi array
        data_masuk = vb.Split(baca_clear, ",")

        'untuk debug serial monitor
        'TextBox2.Text = baca                  '''''''''''''''''''''''''''''''''''''''''' ini di ganti dengan serial yg bacanya banyak, baca ke bawah (Sudah)

        'menghitung (jumlah, koma)
        'jumlah karakter dikurangi jumlah karakter (koma) yang telah di hapus
        jumlah_koma = Len(baca) - Len(Replace(baca, ",", ""))

        If jumlah_koma = 35 Then
            'Data Parameter Atmosfer
            Label28.Text = data_masuk(0) 'ID team'
            Label9.Text = data_masuk(2) 'altitude'
            Label17.Text = data_masuk(3) 'temperatur'
            Label18.Text = data_masuk(4) 'humidity'
            Label16.Text = data_masuk(5) 'pressure'
            Label15.Text = data_masuk(6) 'WC / Arah Angin'
            Label10.Text = data_masuk(7) 'WS / Kecepatan Angin'
            Label41.Text = data_masuk(8) 'arah'
            Label6.Text = data_masuk(9) 'latitude'
            Label5.Text = data_masuk(10) 'longitude'

            'Label untuk Data Parameter Atmosfer
            ID = data_masuk(0) 'ID team'
            Wkt = data_masuk(1) 'Waktu'
            Altitude = data_masuk(2) 'Altitude'
            Temperature = data_masuk(3) 'Temperatur'
            Humidity = data_masuk(4) 'Humidity'
            Pressure = data_masuk(5) 'Pressure'
            WC = data_masuk(6) 'Arah Angin'
            WS = data_masuk(7) 'Kecepatan Angin
            arah = data_masuk(8) 'arah'
            Latitude = data_masuk(9) 'Latitude'
            Longitude = data_masuk(10) 'Longitude'

            'Label untuk Masukan Antenna Tracker
            IDgps = data_masuk(11)
            wkt2 = data_masuk(12)
            Latitude2 = data_masuk(13)
            HeadingLatitude2 = data_masuk(14)
            Longitude2 = data_masuk(15)
            HeadingLongitude2 = data_masuk(16)
            GPSstatus = data_masuk(17)
            'Label31.Text = data_masuk(18)
            Sats = data_masuk(18)
            'Label35.Text = data_masuk(19)
            HDOP = data_masuk(19)
            'Label9.Text = data_masuk(20)
            Altitude = data_masuk(20)
            satuan1 = data_masuk(21)
            AltSL = data_masuk(22)
            satuan2 = data_masuk(23)
            kosong = data_masuk(24)
            CKS = data_masuk(25)
            IDgps2 = data_masuk(26)
            kosong2 = data_masuk(27)
            T = data_masuk(28)
            kosong3 = data_masuk(29)
            M = data_masuk(30)
            GSpeed = data_masuk(31)
            N = data_masuk(32)
            GSpeed2 = data_masuk(33)
            K = data_masuk(34)
            CKS2 = data_masuk(35)
        End If


        TextBox2.AppendText(ID & "," & Wkt & "," & Altitude & "," & Temperature & "," & Humidity & "," & Pressure & "," & WC & "," & WS & "," & arah & "," & Latitude & "," & Longitude & vbCrLf)
        TextBox1.AppendText(IDgps & "," & wkt2 & "," & Latitude2 & "," & HeadingLatitude2 & "," & Longitude2 & "," & HeadingLongitude2 & "," & GPSstatus & "," & Sats & "," & HDOP & "," & Altitude & "," & satuan1 & "," & AltSL & "," & satuan2 & "," & kosong & "," & CKS & vbCrLf & IDgps2 & "," & kosong2 & "," & T & "," & kosong3 & "," & M & "," & GSpeed & "," & N & "," & GSpeed2 & "," & K & "," & CKS2 & vbCrLf)

        payload_marker.Position = New PointLatLng(Latitude, Longitude)
        If data_pertama Then
            GMapControl1.Position = New PointLatLng(Latitude, Longitude)
            data_pertama = False
        End If

        'Tambahan
        If rc6in > tth Then
            marker_layer.Markers.Add(New GMapPoint(New PointLatLng(Latitude, Longitude), 5, rc6in)) ' sesuaikan jari2nya
            dgv1.Rows.Add(New String() {CStr(rc6in), CStr(Latitude), CStr(Longitude)})
        End If
        '------------

        GMapControl1.UpdateMarkerLocalPosition(payload_marker)
        GMapControl1.Invalidate()

        Double.TryParse(WC, WC1)
        Double.TryParse(WS, WS1)

        Chart1.Series("Series1").Points.AddXY(WC1 / 100, WS1 / 100)
        Chart2.Series("Series1").Points.AddXY(WS1 / 100, Altitude)
        Chart3.Series("Series1").Points.AddXY(Pressure, Altitude)
        Chart4.Series("Series1").Points.AddXY(WC1 / 100, Altitude)
        Chart5.Series("Series1").Points.AddXY(Humidity, Altitude)
        Chart6.Series("Series1").Points.AddXY(Temperature, Altitude)

    End Sub

    'Tambahan
    Public Class GMapPoint
        Inherits GMapMarker
        Private point_ As PointLatLng
        Private size_ As Single
        Private val_, x As Double

        Public Property Point() As PointLatLng
            Get
                Return point_
            End Get
            Set
                point_ = Value
            End Set
        End Property
        Public Sub New(p As PointLatLng, size As Integer, val As Double)
            MyBase.New(p)
            point_ = p
            size_ = size
            val_ = val
            x = (val - 1500) / 500 * 100  ' 1500 diganti dg nilai threshold, angka 500 diganti dg (nilai maximum- threshold)
        End Sub

        Public Overrides Sub OnRender(g As Graphics)
            Dim mycolor As Color = Color.FromArgb(255, 255.0F * 1 * (x / 100), 255.0F * 1 * (1 - (x / 100)), 0)
            Dim myBrush As New SolidBrush(mycolor)
            'g.FillRectangle(Brushes.Red, LocalPosition.X, LocalPosition.Y, size_, size_)
            'OR 
            g.FillEllipse(myBrush, LocalPosition.X, LocalPosition.Y, size_, size_)
            'OR whatever you need

        End Sub
    End Class
    '------------

    Public Sub updateTextBox3()
        'TextBox1.AppendText(IDgps & "," & wkt2 & "," & Latitude2 & "," & HeadingLatitude2 & "," & Longitude2 & "," & HeadingLongitude2 & "," & GPSstatus & "," & Sats & "," & HDOP & "," & Altitude & "," & satuan1 & "," & AltSL & "," & satuan2 & "," & kosong & "," & CKS & vbCr & IDgps2 & "," & kosong2 & "," & T & "," & kosong3 & "," & M & "," & GSpeed & "," & N & "," & GSpeed2 & "," & K & "," & CKS2 & vbCrLf)
        SerialPort2.WriteLine(TextBox1.Text)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'variable untuk tanggal dan jam
        Label44.Text = Date.Now.ToString("dd-MMMM-yyyy")
        Label46.Text = Date.Now.ToString("H:mm:ss")

        't1 = 
    End Sub

    'Berbagai macam Button inisialisasi dan fungsi

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Panel1.Visible = True
        Button1.BackgroundImage = My.Resources.button_1
        TableLayoutPanel2.Visible = False
        Panel1.Visible = True
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim streamWrite As New System.IO.StreamWriter("D:\KULIAH\SKRIPSI\DATA LOG\CSV\" & Date.Now.ToString("dd MMMM yyyy") & "_Kombat2020_filter.csv")
        streamWrite.Write(TextBox2.Text)
        streamWrite.Close()
        Chart1.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Diagram Polar Arah Angin.jpg", ChartImageFormat.Jpeg)
        Chart2.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Kecepatan Angin.jpg", ChartImageFormat.Jpeg)
        Chart3.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Tekanan Udara.jpg", ChartImageFormat.Jpeg)
        Chart4.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Arah Angin.jpg", ChartImageFormat.Jpeg)
        Chart5.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Kelembaban Relatif.jpg", ChartImageFormat.Jpeg)
        Chart6.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Temperatur.jpg", ChartImageFormat.Jpeg)
        MessageBox.Show("Payload Information successfull saved to file (D:\KULIAH\SKRIPSI\DATA LOG\CSV\Kombat2020_filter.csv)", "Payload Message")
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click ' ini di tambah convert data Raw, mirror kan "baca" ke label atau textbox (sudah)
        Dim streamWrite As New System.IO.StreamWriter("D:\KULIAH\SKRIPSI\DATA LOG\TXT\" & Date.Now.ToString("dd MMMM yyyy") & "_Kombat2020_filter.txt")
        streamWrite.Write(TextBox2.Text)
        streamWrite.Close()
        MessageBox.Show("Payload Information successfull converted into .txt file (D:\KULIAH\SKRIPSI\DATA LOG\TXT\Kombat2020_filter.txt)", "Payload Message")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Label49.Text = Latitude & "," & Longitude
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Label30.Text = Date.Now.ToString("H:mm:ss")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Chart1.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Diagram Polar Arah Angin.jpg", ChartImageFormat.Jpeg)
        Chart2.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Kecepatan Angin.jpg", ChartImageFormat.Jpeg)
        Chart3.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Tekanan Udara.jpg", ChartImageFormat.Jpeg)
        Chart4.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Arah Angin.jpg", ChartImageFormat.Jpeg)
        Chart5.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Kelembaban Relatif.jpg", ChartImageFormat.Jpeg)
        Chart6.SaveImage("D:\KULIAH\SKRIPSI\DATA LOG\JPEG\" & Date.Now.ToString("dd MMMM yyyy") & "_Grafik Parameter Temperatur.jpg", ChartImageFormat.Jpeg)
        MessageBox.Show("All Grapgh successfull converted into .jpeg, (D:\KULIAH\SKRIPSI\DATA LOG\JPEG)", "Graph Message")
    End Sub

    'Untuk Pemasukan COM Antena Tracker
    Private Sub TrackerButton9_Click(sender As Object, e As EventArgs) Handles TrackerButton9.Click
        SerialPort2.PortName = "COM13" 'ComboBox1.Text ' '<---SerialPort2 PortName membaca isi dari PortComboBox1 (COM XX)'
        SerialPort2.BaudRate = "57600" 'ComboBox2.Text '"57600" '<---SerialPort2 BaudRate membaca isi dari PortComboBox2 (Nilai BaudRate)'
        SerialPort2.DataBits = 8
        SerialPort2.Parity = Parity.None
        SerialPort2.StopBits = StopBits.One
        SerialPort2.Handshake = Handshake.None
        SerialPort2.Encoding = System.Text.Encoding.Default
        SerialPort2.Open()
        Label2.Text = "Connected"
        Label2.ForeColor = Color.LimeGreen
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SerialPort2.Close()
        Label2.Text = "Dis-connected"
        Label2.ForeColor = Color.Red
    End Sub
End Class
