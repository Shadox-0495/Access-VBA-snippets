VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Enum DBType
vbaLocal = 0
vbaExternal = 1
End Enum

Enum ClrVal
vbaNone = 0
vbaWhite = 1
vbaRed = 2
vbaGreen = 3
vbaBlue = 4
vbaYellow = 5
vbaBlack = 6
End Enum

Private cmStatus As Boolean
Private JDEConStat As Boolean
Private JDEDsn As String
Private JDEUsr As String
Private JDEPass As String

Private Type UfObj
db As DAO.Database
jde As Object
JDErst As Object
UF As Object
Ui As Object
End Type

Private this As UfObj

Public Property Let MsADB(MsDB As DAO.Database)
Set this.db = MsDB
End Property

Public Property Get MsADB() As DAO.Database
Set MsADB = this.db
End Property

Public Property Let JDEcon(JDE_con As Object)
Set this.jde = JDE_con
End Property

Public Property Get JDEcon() As Object
Set JDEcon = this.jde
End Property

Public Property Let UsrForm(sUsrForm As Object)
Set this.UF = sUsrForm
Set this.Ui = this.UF.Controls
End Property

Public Property Get UsrForm() As Object
Set UsrForm = this.UF
End Property

Public Function GetConStatus()
GetConStatus = cmStatus
End Function

Public Function C_JDE_Auto(EstacionID As Integer) As Boolean
Dim conectado As Boolean: conectado = False 'Variable de validacion para confirmar el proceso

'Abrir recordset para recuperar todos los usuarios de JDE y contraseñas asignados al AreaID de la EstacionID
Dim rst As DAO.Recordset
Set rst = this.db.OpenRecordset("SELECT t2.ID,t2.Usuario,t2.Contraseña FROM Estaciones t1 INNER JOIN [Usuarios JDE] t2 ON t2.AreaID=t1.AreaID WHERE t1.ID=" & EstacionID & " AND t2.Intentos<3")

If rst.EOF Then GoTo fin 'Si el recordset esta vacio, ir al final del proceso

'Si el recordset no esta vacio, ir probando cada usuario y contraseña retornados para hcer la conexion a JDE
Do Until rst.EOF

    If Not (this.jde Is Nothing) Then 'Verificar si la variable de conexion esta siendo utilizada
        If this.jde.State = 1 Then GoTo fin 'Verificar la conexion si esta abierta e ir al final del proceso -> variable de conexion creada correctamente
    End If
    On Error GoTo ErrHandler 'Si hay algun error en el siguiente proceso, ir al ErrHandler
        conectado = Conectar_JDE(rst.Fields("Usuario"), rst.Fields("Contraseña")) 'Llamar al proceso que crea la conexion a JDE
        AccessSQL "UPDATE [Usuarios JDE] SET Intentos=0 WHERE ID=" & rst.Fields("ID") 'Si se logro establecer conexion a JDE, reiniciar los intentos
        GoTo siguiente
ErrHandler:
    AccessSQL "UPDATE [Usuarios JDE] SET Intentos=Intentos+1 WHERE ID=" & rst.Fields("ID") 'Si en el proceso de conexion a JDE se encontro un error, aumentar los intentos
    'Registrar los errores en la tabla de errores de usuarios JDE
    AccessSQL "INSERT INTO [Usuarios JDE errores] (Fecha,UsuarioJDEID,Error,Maquina,Usuario) VALUES('" & Date & "'," & rst.Fields("ID") & ",'" & Err.Description & "','" & GetPCName & "','" & GetWinUser & "')"
    'GoTo fin
siguiente:
    rst.MoveNext 'Seguir con el siguiente usuario del recordset
Loop
fin:
rst.Close: Set rst = Nothing
C_JDE_Auto = conectado 'Retornar variable booleana que verifica si se logro conectar a JDE
End Function

Public Function Conectar_JDE(Usuario As String, contra As String) As Boolean
Dim conectado As Boolean: conectado = False 'Variable de validacion para confirmar la conexion a JDE

If this.jde Is Nothing Then Set this.jde = CreateObject("ADODB.Connection") 'Si la variable de conexion esta "Limpia" asignarle un Objecto de tipo "ADODB"
Dim cs As String 'Variable para la cadena de conexion

'Cadena de conexion a JDE utilizando DSN ->JDE LIVE
'cs = "DSN=JDE LIVE;DRIVER=Client Access ODBC Driver (64-bit); " & _
'                           ";UID =" & Usuario & _
'                           ";PWD =" & Contra

'Cadena de conexion a JDE sin utilizar DSN -> JDE Live, NO MODIFICAR
cs = "DRIVER=iSeries Access ODBC Driver;"
cs = cs & "AllowDataCompression = 1;AllowProcCalls = 0;AllowUnsupportedChar = 0;BlockFetch = 1;BlockSizeKB = 256;CatalogOptions = 3;CCSID = 0;CommitMode = 2;Concurrency = 0;ConcurrentAccessResolution = 0;ConnectionType = 0;ConvertDateTimeToChar = 0;CursorSensitivity = 0;DateFormat = 5;DateSeparator = 1;"
cs = cs & "DB2SQLStates = 0;Debug = 0;DecfloatErrorOption = 0;DecFloatRoundMode = 0;Decimal = 0;DefaultLibraries = JDLDTA JDLDTA;DefaultPackage = QGPL/DEFAULT(IBM),2,0,1,0,512;DefaultPkgLibrary = QGPL;DelimitNames = 0;ExtendedColInfo = 0;ExtendedDynamic = 0;ExtendedType = 0;ForceTranslation = 1;"
cs = cs & "Graphic = 0;HexParserOpt = 0;LanguageID = ENU;LazyClose = 0;LibraryView = 0;MapDecimalFloatDescribe = 1;MaxFieldLength = 32;MaximumDecimalPrecision = 31;MaximumDecimalScale = 31;MinimumDivideScale = 0;Naming = 0;ODBCRemarks = 0;PreFetch = 1;QueryOptimizeGoal = 0;QueryStorageLimit = -1;"
cs = cs & "QueryTimeout = 1;ReportDataTypes = 1;SearchPattern = 1;Signon = 3;SortSequence = 0;SortWeight = 0;SQLConnectPromptMode = 0;SSL = 2;System = S101B1CA;TimeFormat = 0;TimeSeparator = 0;Trace = 0;TranslationOption = 0;TrueAutoCommit = 0;UnicodeSQL = 0;XALockTimeout = 0;XALooselyCoupledSupport = 1;"
cs = cs & "XATransactionTimeout = 0;XMLCurrentImplicitParse = 0;XMLDeclarationFormat = 7;"
cs = cs & "UID=" & Usuario & ";" '->Nombre del usuario de JDE
cs = cs & "PWD=" & contra '->Contraseña del usuario de JDE
                           
'Si anteriormente ya existia la variable de conexion y esta en uso, cerrarla
If Not (this.jde Is Nothing) Then
    If this.jde.State = 1 Then this.jde.Close
End If

this.jde.Open (cs) 'Abrir la variable de conexion con el nuevo usuario y contraseña
conectado = True 'El proceso de conexion finalizo correctamente, cambiar conexion de Falso a Verdadero

Conectar_JDE = conectado 'Retornar variable booleana que verifica si se logro conectar a JDE
End Function

Public Sub Desconectar_JDE()
If this.jde Is Nothing Then Exit Sub 'Si no hay conexion a JDE, salir del proceso
    'estado:
    '0=cerrado
    '1=abierto
    '2=conectando
    '3=ejecutando
    '4=recuperando
If this.jde.State = 1 Then this.jde.Close 'Si hay conexion a JDE y esta abierta, cerar la conexion
Set this.jde = Nothing '"Limpiar" la variable de la conexion
End Sub

Public Function Confirmar_conexion_JDE(EstacionID As Integer) As Boolean
Dim val As Boolean: val = False
'If this.Pfrm.tbInventario Then GoTo fin 'Verificar si el formulario esta en modo de inventario
If this.jde Is Nothing Then val = C_JDE_Auto(EstacionID) 'verificar la conexion a JDE, si no esta establecida intentara abrirla
If this.jde Is Nothing Then GoTo fin
If this.jde.State = 0 Then GoTo fin 'Si aun la conexion no esta abierta, saltar el proceso
val = True
fin:
    Confirmar_conexion_JDE = val 'Retornar la comprobacion del estado de la conexion
End Function

Public Function Confirmar_personal(CargoID As Integer, Codigo As String)
Confirmar_personal = AccessSQL("SELECT Nombre FROM Personal WHERE CargoID=" & CargoID & " AND Codigo='" & Codigo & "'")
End Function

Public Sub Descargar_Informacion_lote_JDE(lote As String, LoginID As Integer, Optional TblSubform As Object)
Set this.JDErst = CreateObject("ADODB.Recordset")
Dim sql As String: sql = ""

AccessSQL "DELETE * FROM JDE_consulta WHERE LoginID=" & LoginID & " AND InfoID=1"

If Not TblSubform Is Nothing Then TblSubform.Requery

'-----------------Ubicacion actual del lote en JDE, segun tabla 41021
sql = "SELECT tt1.*,t4.IMDSC1 AS Desc1,t4.IMDSC2 AS Desc2 FROM ( "
sql = sql & " SELECT t1.LIUPMJ AS Fecha,t1.LITDAY AS Hora,TRIM(t1.LIMCU) AS Planta,TRIM(t1.LILOCN) AS Locacion,CASE WHEN t2.ILKIT IS NOT NULL THEN t2.ILKIT ELSE t1.LIITM END AS SKU,t2.ILDOC AS Orden,CAST(t1.LIPQOH/100 AS DOUBLE) AS Kg,t3.UMCONV AS Fc,TRIM(t1.LILOTS) AS Estado"
sql = sql & " FROM ( F41021 AS t1 LEFT JOIN F4111 AS t2 ON (t1.LILOTN=t2.ILLOTN AND t2.ILLOCN=t1.LILOCN AND t2.ILMMCU=t1.LIMCU)) LEFT JOIN F41002 t3 ON (t3.UMMCU=t1.LIMCU AND t3.UMITM=t1.LIITM AND t3.UMUM='KG' AND t3.UMRUM='DZ')"
sql = sql & " WHERE t1.LILOTN = '" & lote & "' AND (t2.ILDOC>31114 OR t1.LILOCN='FNLT01' OR t1.LILOCN='FNLT02' OR t1.LILOCN='DYLT01' OR t1.LILOCN='HOLDGS')"
sql = sql & " GROUP BY t1.LIITM,t2.ILKIT, t1.LILOCN, t1.LIPQOH, t2.ILDOC, t1.LIUPMJ, t1.LITDAY, t1.LIMCU,t3.UMCONV,t1.LILOTS"
sql = sql & " ORDER BY t1.LIUPMJ DESC, t1.LITDAY DESC"
sql = sql & " ) AS tt1 LEFT JOIN F4101 AS t4 ON t4.IMITM=tt1.SKU"

this.JDErst.Open sql, this.jde
If this.JDErst.EOF Then GoTo transacciones

Do Until this.JDErst.EOF
    sql = "INSERT INTO JDE_consulta(LoginID,InfoID,Fecha,Hora,Planta,Locacion,SKU,Orden,Kg,FC,Desc1,Desc2,Estado)" & _
                             " VALUES(" & LoginID & ",1," & this.JDErst.Fields("Fecha") & "," & this.JDErst.Fields("Hora") & ",'" & this.JDErst.Fields("Planta") & "','" & this.JDErst.Fields("Locacion") & "'," & this.JDErst.Fields("SKU") & "," & Nz(this.JDErst.Fields("Orden"), 0) & "," & this.JDErst.Fields("Kg") & "," & IIf(this.JDErst.Fields("FC") > 0, Round(1 / (this.JDErst.Fields("FC") / 10000000), 4), 0) & ",'" & Trim(this.JDErst.Fields("Desc1")) & "','" & Trim(this.JDErst.Fields("Desc2")) & "','" & this.JDErst.Fields("Estado") & "')"
    Debug.Print sql
    AccessSQL sql
this.JDErst.MoveNext
Loop

GoTo cerrar_rst

transacciones:
'-----------------Ubicacion actual del lote en JDE, segun tabla 4111, Order por fecha y hora de mayor a menor, tomar el primer registro como la ubicacion actual y luego modificar todos los demas a 0
sql = "SELECT tt1.*,t3.IMDSC1 AS Desc1,t3.IMDSC2 AS Desc2 FROM ("
sql = sql & "SELECT t1.ILCRDJ AS Fecha,t1.ILTDAY AS Hora,TRIM(t1.ILMCU) AS Planta,TRIM(t1.ILLOCN) AS Locacion,CASE WHEN t1.ILKIT IS NOT NULL AND t1.ILKIT>0 THEN t1.ILKIT ELSE t1.ILITM END AS SKU,t1.ILDOC AS Orden,CAST(t1.ILTRQT/100 AS Double) AS Kg,t1.ILLOTS AS Estado,t2.UMCONV AS Fc"
sql = sql & " FROM F4111 t1 LEFT JOIN F41002 t2 ON (t2.UMMCU=t1.ILMCU AND t2.UMITM=t1.ILITM AND t2.UMUM='KG' AND t2.UMRUM='DZ')"
sql = sql & " WHERE t1.ILLOTN='" & lote & "' AND t1.ILDOC>31114 AND t1.ILTRQT>0"
sql = sql & " GROUP BY  t1.ILCRDJ,t1.ILTDAY,t1.ILMCU,t1.ILLOCN,t1.ILKIT,t1.ILITM,t1.ILDOC,t1.ILTRQT,t1.ILLOTS,t2.UMCONV"
sql = sql & " ORDER BY t1.ILCRDJ DESC,t1.ILTDAY DESC"
sql = sql & ") AS tt1 LEFT JOIN F4101 t3 ON t3.IMITM=tt1.SKU"

this.JDErst.Close: this.JDErst.Open sql, this.jde
If this.JDErst.EOF Then GoTo cerrar_rst
Dim tmp_fecha, tmp_hr, tmp_conteo As Long
Dim tmp_kg, tmp_kg_anterior As Double: tmp_kg = this.JDErst.Fields("Kg"): tmp_kg_anterior = this.JDErst.Fields("Kg")
tmp_conteo = 0
tmp_fecha = this.JDErst.Fields("Fecha")
tmp_hr = this.JDErst.Fields("Hora")
Do Until this.JDErst.EOF
    If tmp_conteo > 0 Then
        If tmp_kg_anterior <> this.JDErst.Fields("Kg") Then tmp_kg = this.JDErst.Fields("Kg") Else tmp_kg = 0
    End If
    sql = "INSERT INTO JDE_consulta(LoginID,InfoID,Fecha,Hora,Planta,Locacion,SKU,Orden,Kg,FC,Desc1,Desc2,Estado)" & _
                             " VALUES(" & LoginID & ",1," & this.JDErst.Fields("Fecha") & "," & this.JDErst.Fields("Hora") & ",'" & this.JDErst.Fields("Planta") & "','" & this.JDErst.Fields("Locacion") & "'," & this.JDErst.Fields("SKU") & "," & Nz(this.JDErst.Fields("Orden"), 0) & "," & tmp_kg & "," & IIf(this.JDErst.Fields("FC") > 0, Round(1 / (this.JDErst.Fields("FC") / 10000000), 4), 0) & ",'" & Trim(this.JDErst.Fields("Desc1")) & "','" & Trim(this.JDErst.Fields("Desc2")) & "','" & this.JDErst.Fields("Estado") & "')"
    AccessSQL sql
    tmp_conteo = tmp_conteo + 1
this.JDErst.MoveNext
Loop
'AccessSQL "UPDATE JDE_consulta SET Kg=0 WHERE Fecha<>" & tmp_fecha & " AND Hora<>" & tmp_hr & " AND LoginID=" & LoginID & " AND InfoID=1"

cerrar_rst:
this.JDErst.Close: Set this.JDErst = Nothing
End Sub

Public Sub Descargar_Informacion_JDE_GSKU(GSKU As Long, LoginID As Integer)
Set this.JDErst = CreateObject("ADODB.Recordset")
AccessSQL "DELETE * FROM JDE_consulta WHERE LoginID=" & LoginID & " AND InfoID=2"
Dim sql As String: sql = ""
'sql = "SELECT t1.IXKIT AS DSKU,t2.UMCONV AS FC,SUM(t3.WAUORG) AS Kg FROM " & _
'                         "(" & _
'                         "F3002 t1 LEFT JOIN F41002 t2 ON (t2.UMITM=t1.IXITM AND t2.UMMCU=t1.IXMMCU AND TRIM(t2.UMUM)='KG' AND TRIM(t2.UMRUM)='DZ')" & _
'                         ") LEFT JOIN F4801L13 t3 ON t3.WAITM=t1.IXITM " & _
'                         " WHERE t1.IXITM=" & GSKU & " AND (t3.WASRST>='05' AND t3.WASRST<='40') GROUP BY t1.IXKIT,t2.UMCONV"
                         
sql = "SELECT t1.IXKIT AS DSKU,t2.UMCONV AS FC,SUM(t3.WAUORG) AS Kg FROM " & _
                         "(" & _
                         "F3002 t1 LEFT JOIN F41002 t2 ON (t2.UMITM=t1.IXITM AND t2.UMMCU=t1.IXMMCU AND TRIM(t2.UMUM)='KG' AND TRIM(t2.UMRUM)='DZ')" & _
                         ") LEFT JOIN F4801L13 t3 ON (t3.WAITM=t1.IXITM AND (t3.WASRST>='05' AND t3.WASRST<='40')) " & _
                         " WHERE t1.IXITM=" & GSKU & " GROUP BY t1.IXKIT,t2.UMCONV"
                         
this.JDErst.Open sql, this.jde
If this.JDErst.EOF Then GoTo cerrar_rst
sql = "INSERT INTO JDE_consulta(LoginID,InfoID,SKU,Kg,FC) VALUES(" & LoginID & ",2," & this.JDErst.Fields("DSKU") & "," & IIf(IsNull(this.JDErst.Fields("Kg")), 0, Round(((this.JDErst.Fields("Kg") / 100) / 90), 0)) & "," & IIf(this.JDErst.Fields("FC") > 0, Round(1 / (this.JDErst.Fields("FC") / 10000000), 4), 0) & ")"
Debug.Print sql
AccessSQL sql

cerrar_rst:
this.JDErst.Close: Set this.JDErst = Nothing
End Sub

Public Sub Descargar_Informacion_JDE_DSKU(DSKU As Long, LoginID As Integer)
Dim sql As String: sql = ""
Set this.JDErst = CreateObject("ADODB.Recordset")
AccessSQL "DELETE * FROM JDE_consulta WHERE LoginID=" & LoginID & " AND InfoID=2"

sql = "SELECT WADOCO AS Orden, WASRST AS Estado,(WAUORG-WASOQS)/100 AS Kg,WADRQJ AS Fecha FROM F4801 WHERE WAITM=" & DSKU & " AND TRIM(WAMCU)='RIOHS' AND (WALOCN='DYLT01' OR WALOCN='DYTR01') AND CAST(WASRST AS Int) BETWEEN 26 AND 75 ORDER BY WADRQJ"
this.JDErst.Open sql, this.jde
If this.JDErst.EOF Then GoTo cerrar_rst

Do Until this.JDErst.EOF
    sql = "INSERT INTO JDE_consulta(LoginID,InfoID,Fecha,Orden,Kg,Estado)" & _
                             " VALUES(" & LoginID & ",2," & this.JDErst.Fields("Fecha") & "," & Nz(this.JDErst.Fields("Orden"), 0) & "," & this.JDErst.Fields("Kg") & ",'" & this.JDErst.Fields("Estado") & "')"
    AccessSQL sql
this.JDErst.MoveNext
Loop

cerrar_rst:
this.JDErst.Close: Set this.JDErst = Nothing
End Sub

Public Sub Descargar_Informacion_JDE_WO(WO As Long, LoginID As Integer, SbForm As Object)
Dim sql As String: sql = ""
AccessSQL "DELETE * FROM JDE_consulta WHERE LoginID=" & LoginID & " AND InfoID=2": SbForm.Requery
sql = "SELECT WAITM AS SKU, CAST(WASRST AS INT) AS Estado, TRIM(WATBM) AS Bt FROM F4801L13 WHERE WADOCO=" & WO

Set this.JDErst = CreateObject("ADODB.Recordset")
this.JDErst.Open sql, this.jde
If this.JDErst.EOF Then GoTo cerrar_rst

Do Until this.JDErst.EOF
    sql = "INSERT INTO JDE_consulta(LoginID,InfoID,SKU,Estado)" & _
                             " VALUES(" & LoginID & ",2," & this.JDErst.Fields("SKU") & ",'" & this.JDErst.Fields("Estado") & "')"
    AccessSQL sql
this.JDErst.MoveNext
Loop

cerrar_rst:
this.JDErst.Close: Set this.JDErst = Nothing
End Sub


Public Sub AccessQuit()
Application.Quit
End Sub

Function SwicthAppPrinter(sPrinterName As String) As Boolean
    On Error GoTo Error_Handler
    Dim prtAvailPrinters As Printer
 
    For Each prtAvailPrinters In Application.Printers
        If prtAvailPrinters.DeviceName = sPrinterName Then
            Set Application.Printer = prtAvailPrinters
            SwicthAppPrinter = True
            Exit For
        End If
    Next prtAvailPrinters
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: SwicthAppPrinter" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
                Exit For
            Case Else
                IsLetter = False
        End Select
    Next
End Function

Public Function GetWinUser() As String
    GetWinUser = CreateObject("WScript.Network").UserName
End Function

Public Function GetPCName() As String
GetPCName = Environ$("computername")
End Function

Public Function GetPcIP()
Dim myWMI As Object, myobj As Object, itm
Set myWMI = GetObject("winmgmts:\\.\root\cimv2")
Set myobj = myWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each itm In myobj
  GetPcIP = itm.IPAddress(0)
  Exit Function
Next
End Function

Public Function AccessSQL(sql As Variant) As Variant
Dim tmp As Variant: tmp = Null
If this.db Is Nothing Then GoTo fin
If (InStr(sql, "DELETE") Or InStr(sql, "delete") Or InStr(sql, "Delete")) > 0 Or (InStr(sql, "UPDATE") Or InStr(sql, "update") Or InStr(sql, "Update")) > 0 Or (InStr(sql, "INSERT") Or InStr(sql, "Insert") Or InStr(sql, "insert")) > 0 Then this.db.Execute sql, dbFailOnError: GoTo fin

Dim rst As DAO.Recordset
Set rst = this.db.OpenRecordset(sql)
If Not rst.EOF Then tmp = rst.Fields(0)
rst.Close: Set rst = Nothing
fin:
AccessSQL = tmp
End Function

Public Function ExcelSQLSelect(sql As String, File As String)
Dim ConString As String
Dim ADOcn, ADOrs As Object
Set ADOcn = CreateObject("ADODB.Connection")
Set ADOrs = CreateObject("ADODB.Recordset")
On Error GoTo msg
ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & File & "';Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;"";"
ADOcn.Open ConString
ADOrs.Open sql, ADOcn, 3, 3, &H1
ExcelSQLSelect = ADOrs(0)
msg:
MsgBox ("No se pudo establecer coneccion al archivo Excel")
End
End Function

Public Function ChkValue(UserForm As Object, ItemName As String) As Boolean
Dim chk As Boolean: chk = False
If Len(ItemName) = 0 Then GoTo fin

chk = True
If UBound(Split(ItemName, ";")) = 0 Then ItemName = ItemName & ";"

For Each Item In Split(ItemName, ";")
    If Len(Item) = 0 Then GoTo siguiente
    Select Case UserForm.Controls(Item).ControlType
    Case acLabel, acCommandButton
        If Len(UserForm.Controls(Item).Caption) = 0 Then chk = False: GoTo fin
    Case acTextBox, acComboBox
        If Nz(UserForm.Controls(Item).Value, "") = "" Then chk = False: FormItemsOpt UserForm, CStr(Item), , , vbaWhite, vbaBlack, vbaRed, 1, 3: GoTo fin Else FormItemsOpt UserForm, CStr(Item), , , vbaWhite, vbaBlack, vbaBlack, 1, 1
    End Select
siguiente:
Next Item

fin:
ChkValue = chk
End Function

Public Sub FormItemsAsignVal(UserForm As Object, Item As String, val As Variant)
If Len(Item) = 0 Then Exit Sub
If UBound(Split(Item, ";")) = 0 Then Item = Item & ";"

For Each itm In Split(Item, ";")
If Len(itm) = 0 Then GoTo siguiente
    Select Case UserForm.Controls(itm).ControlType
        Case acLabel, acCommandButton
            UserForm.Controls(itm).Caption = Nz(val, "")
        Case acTextBox, acComboBox, acToggleButton, acCheckBox, acOptionButton
            UserForm.Controls(itm).Value = val
    End Select
siguiente:
Next itm
End Sub

Public Sub FormItemsOpt(UserForm As Object, Item As String, Optional Visible As Boolean = True, Optional Enabled As Boolean = True, Optional BkColor As ClrVal = 0, Optional FrColor As ClrVal = 0, Optional BrColor As ClrVal = 0, Optional BrStyle As Integer = 0, Optional BrWidth As Integer = 0)
If Len(Item) = 0 Then Exit Sub
If UBound(Split(Item, ";")) = 0 Then Item = Item & ";"

For Each itm In Split(Item, ";")
    If Len(itm) = 0 Then GoTo siguiente
    Select Case UserForm.Controls(itm).ControlType
        Case acLabel
            UserForm.Controls(itm).Visible = Visible
        Case acTextBox, acComboBox, acCommandButton, acOptionGroup, acObjectFrame, acListBox, 112
            UserForm.Controls(itm).Visible = Visible
            UserForm.Controls(itm).Enabled = Enabled
    End Select
        
    If BkColor <> vbaNone Then UserForm.Controls(itm).BackColor = vbaColors(BkColor)
    If FrColor <> vbaNone Then UserForm.Controls(itm).ForeColor = vbaColors(FrColor)
    If BrColor <> vbaNone Then UserForm.Controls(itm).BorderColor = vbaColors(BrColor)
    If BrStyle <> 0 Then UserForm.Controls(itm).BorderStyle = BrStyle
    If BrWidth <> 0 Then UserForm.Controls(itm).BorderWidth = BrWidth
siguiente:
Next itm
End Sub

Private Function vbaColors(ClrName As ClrVal)
If ClrName <> vbaNone Then
If ClrName = vbaWhite Then vbaColors = RGB(255, 255, 255)
If ClrName = vbaBlack Then vbaColors = RGB(0, 0, 0)
If ClrName = vbaRed Then vbaColors = RGB(255, 0, 0)
If ClrName = vbaGreen Then vbaColors = RGB(0, 255, 0)
If ClrName = vbaBlue Then vbaColors = RGB(0, 0, 255)
If ClrName = vbaYellow Then vbaColors = RGB(255, 255, 0)
End If
End Function

Public Sub FormChange(Optional NewFormName As Variant = "", Optional NewFormArgs As Variant = "")
DoCmd.Close
If NewFormName <> "" Then DoCmd.OpenForm FormName:=NewFormName, OpenArgs:=NewFormArgs
End Sub

Public Sub ClearList(UserForm As String, Item As String)
If Item <> "" And UBound(Split(Item, ";")) <> -1 Then
Dim ItmArray() As String
ItmArray = Split(Item, ";")
 For I = 0 To UBound(ItmArray)
  For j = Forms(UserForm).Controls(ItmArray(I)).ListCount - 1 To 0 Step -1
   Forms(UserForm).Controls(ItmArray(I)).RemoveItem (j)
  Next
 Next
Else
 For I = Forms(UserForm).Controls(Item).ListCount - 1 To 0 Step -1
  Forms(UserForm).Controls(Item).RemoveItem (I)
 Next
End If

End Sub


Public Sub MakeList(UserForm As String, Item As String, Data As Variant)

Dim LbRow As String
Dim DataRow() As String

Select Case VarType(Data)
Case 9

Do Until Data.EOF
 For I = 0 To Data.Fields.Count - 1
 LbRow = LbRow & Data.Fields(I) & IIf(I = Data.Fields.Count - 1, "", ";")
 Next
 Forms(UserForm).Controls(Item).AddItem LbRow
 LbRow = ""
 Data.MoveNext
Loop

Case 8
 If Len(Data) > 0 Then
 DataRow = Split(Data, ",")
 
 If UBound(DataRow) > 0 Then
  For I = 0 To UBound(DataRow)
   Forms(UserForm).Controls(Item).AddItem DataRow(I)
  Next
 Else
  Forms(UserForm).Controls(Item).AddItem Data
 End If
  
 End If
End Select
End Sub

Public Sub MakeListTable(UserForm As String, lbName As String, Data As Variant, Optional Titulo As Integer = 0)
If (Data.Fields.Count - 1) > Forms(UserForm).Controls(lbName).ColumnCount Then MsgBox "El numero de columnas del recordset es mayor que la lista, agregar mas columnas a la lista" & vbNewLine & "Lista: " & Forms(UserForm).Controls(lbName).ColumnCount & " -  Recordset: " & (Data.Fields.Count - 1), vbCritical: GoTo fin

Forms(UserForm).Controls(lbName).AddItem ""
Dim Fila As Long: Fila = Forms(UserForm).Controls(lbName).ListCount - 1
Do Until Data.EOF
    For I = 0 To Data.Fields.Count - 1
        If Fila = 0 And Titulo = 1 Then lbSetData UserForm, lbName, Data.Field(I).Name, CInt(I), CInt(Fila) Else lbSetData UserForm, lbName, Data.Fields(I), CInt(I), CInt(Fila)
    Next
    Forms(UserForm).Controls(lbName).AddItem ""
    Fila = Forms(UserForm).Controls(lbName).ListCount - 1
    Data.MoveNext
Loop

fin:
End Sub

Public Function LbSlctItem(UserForm As String, lbName As String, ColumnNum As Integer)
LbSlctItem = Forms(UserForm).Controls(lbName).Column(ColumnNum, Forms(UserForm).Controls(lbName).ListIndex)
End Function

Public Sub lbDltSlctItem(UserForm As String, lbName As String)
Forms(UserForm).Controls(lbName).RemoveItem (Forms(UserForm).Controls(lbName).ListIndex)
End Sub

Public Function lbRtnVal(UserForm As String, listbox_nombre As String, Optional Columna As Integer, Optional Fila As Integer) As String
lbRtnVal = Forms(UserForm).Controls(listbox_nombre).Column(Columna, Fila)
End Function

Public Sub lbUpdVal(UserForm As String, listbox_nombre As String, valor As Variant, Columna As Integer, Fila As Integer)
Forms(UserForm).Controls(listbox_nombre).Column(Columna, Fila) = valor
End Sub

Public Sub Listview_Table(UserForm As String, Listview_name As String, Data As Variant, Optional Titulo As Integer = 0, Optional Font_name As String = "Arial", Optional Font_size As Integer = 12)

Forms(UserForm).Controls(Listview_name).Font.Name = Font_name
Forms(UserForm).Controls(Listview_name).Font.Size = Font_size

Forms(UserForm).Controls(Listview_name).ListItems.Clear
Forms(UserForm).Controls(Listview_name).ColumnHeaders.Clear

'If Checboxes = 1 Then Forms(UserForm).Controls(Listview_name).ColumnHeaders.Add , , "", 500, lvwColumnLeft

For I = 0 To Data.Fields.Count - 1
    Forms(UserForm).Controls(Listview_name).ColumnHeaders.Add , , IIf(Titulo = 1, Data.Fields(I).Name, ""), 2000, IIf(I = 0, lvwColumnLeft, lvwColumnCenter)
    'If Checboxes = 0 Then
    'If Checboxes = 1 Then Forms(UserForm).Controls(Listview_name).ColumnHeaders.Add , , IIf(Titulo = 1, Data.Fields(i).Name, ""), 2000, lvwColumnCenter
Next

Dim list As ListItem
'Dim checkbx As New CheckBox

Do Until Data.EOF
    Set list = Forms(UserForm).Controls(Listview_name).ListItems.Add(, , Data.Fields(0))
    If (Data.Fields.Count - 1) >= 1 Then
        For I = 1 To Data.Fields.Count - 1
            list.SubItems(I) = Data.Fields(I)
            'If i = (Data.Fields.Count - 1) Then list.SubItems(i).Controls.Add checkbx
        Next
    End If
    Data.MoveNext
Loop

End Sub

Public Function Listview_Table_gData(UserForm As String, Listview_name As String, Optional ByVal Fila As Integer = 1, Optional ByVal Columna As Integer = 0) As String
'lvData.ListItems(1).SubItems(1) = "Hola Mundo"
'lvData.ListItems(1) = "Hola Mundo"
'lvData.ColumnHeaders(2).Width = 5
'Debug.Print lvData.ListItems(1).SubItems(1)
If Columna = 0 Then Listview_Table_gData = Forms(UserForm).Controls(Listview_name).ListItems(Fila) Else Listview_Table_gData = Forms(UserForm).Controls(Listview_name).ListItems(Fila).SubItems(Columna)
End Function

Public Sub Listview_Table_sData(UserForm As String, Listview_name As String, valor As String, Optional ByVal Fila As Integer = 1, Optional ByVal Columna As Integer = 0)
If Columna = 0 Then Forms(UserForm).Controls(Listview_name).ListItems(Fila) = valor Else Forms(UserForm).Controls(Listview_name).ListItems(Fila).SubItems(Columna) = valor
End Sub

Function SelectFile(MultFile As Boolean)
Dim fd As Object
Set fd = Application.FileDialog(3)
fd.AllowMultiSelect = MultFile

If MultFile = False Then
If fd.Show = True Then SelectFile = fd.SelectedItems(1)
Else
If fd.Show = True Then
 If fd.SelectedItems(1) <> vbNullString Then
 Dim TotalFiles, FileCount As Long
 Dim Files() As String
   For Each vrtSelectedItem In fd.SelectedItems
   TotalFiles = TotalFiles + 1
   Next vrtSelectedItem
   ReDim Files(TotalFiles)
   For Each vrtSelectedItem In fd.SelectedItems
   FileCount = FileCount + 1
   Files(FileCount) = vrtSelectedItem
   Next vrtSelectedItem
   SelectFile = Files
  End If
Else
  End
End If
End If

Set fd = Nothing
End Function

Public Sub MkBackup(File As String, NewPath As String, Optional NewName As String = "", Optional ReplaceFile As Boolean = False)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Call fso.CopyFile(File, CStr(IIf(NewName = "", NewPath & "\" & Dir(File), NewPath & "\" & NewName)), ReplaceFile)
Set fso = Nothing
End Sub

Public Sub AccessQueryMod(QueryName As String, sql As String)
CurrentDb.QueryDefs(QueryName).sql = sql
End Sub

Public Sub CreateOutlookMail(MailSibject As String, MailTo As String, Content As String, Optional Attach As Variant = "", Optional Display As Boolean = False, Optional CopyTo As String = "", Optional Important As Integer = 0)
Dim oApp As Object: Set oApp = CreateObject("Outlook.Application")
Dim oCorreo As Object: Set oCorreo = oApp.CreateItem(0)

With oCorreo
 .To = MailTo
 .Cc = CopyTo
 .Subject = MailSibject
 .HTMLBody = Content
     On Error GoTo Continuar
     If UBound(Attach) <> -1 Then
      For I = 0 To UBound(Attach)
       .Attachments.Add Attach(I)
      Next
     End If
Continuar:
 If Display = False Then .Send Else .Display
 If Important > 0 Then .Importance = 2
 
End With
On Error GoTo 0

Set oCorreo = Nothing: Set oApp = Nothing
End Sub

Function Linkear_tablas(bShowSys As Boolean) As String
On Error GoTo Linkear_tablas_Error
    Dim db As DAO.Database
    Dim td As DAO.TableDefs
    Dim Ruta As String: Ruta = AccessSQL("SELECT Ruta FROM Configuracion WHERE ID=1", vbaLocal)
    Set db = DBEngine.Workspaces(0).OpenDatabase(Ruta)
    Set td = db.TableDefs
    For Each t In td    'loop through all the fields of the tables
        If Left(t.Name, 4) = "MSys" And bShowSys = False Then GoTo Continue
        DoCmd.TransferDatabase acLink, "Microsoft Access", Ruta, acTable, t.Name, t.Name
Continue:
    Next
 
    Set td = Nothing
    Set db = Nothing
If Err.Number = 0 Then Exit Function
 
Linkear_tablas_Error:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: listTable" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function

Function Eliminar_tablas_linkeadas()
On Error GoTo Error_Handler
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
 
    DoCmd.SetWarnings False
    Set db = CurrentDb()
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If InStr(CurrentDb.TableDefs(tdf.Name).Connect, "ODBC") = 0 Then DoCmd.DeleteObject acTable, tdf.Name
        End If
    Next
 
Error_Handler_Exit:
    DoCmd.SetWarnings True
    Set tdf = Nothing
    Set db = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Eliminar_tablas_linkeadas" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function
