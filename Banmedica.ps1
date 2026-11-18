#requires -Version 5.1
# Artikos / Banmedica - preparacion LOCAL exclusivamente. Sin envios.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Data
$script:root = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:auditPath = Join-Path $script:root 'Logs\auditoria_procesos.jsonl'
$script:sessionId = [guid]::NewGuid().ToString()
[void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $script:auditPath))
$script:work = Join-Path $env:LOCALAPPDATA 'Artikos_Banmedica\Trabajo'
$script:inbox = Join-Path $env:LOCALAPPDATA 'Artikos_Banmedica\Distribucion'
[void](New-Item -ItemType Directory -Force -Path $script:work,$script:inbox)
$script:settingsFile = Join-Path $env:LOCALAPPDATA 'Artikos_Banmedica\preferencias.json'
$script:csvFolder = Join-Path $env:LOCALAPPDATA 'Artikos_Banmedica\Reportes'
$script:excelFolder = $script:csvFolder
if(Test-Path -LiteralPath $script:settingsFile){
 try{
  $saved=Get-Content -LiteralPath $script:settingsFile -Raw | ConvertFrom-Json
  if($saved.Distribucion){$script:inbox=[string]$saved.Distribucion}
  if($saved.Trabajo){$script:work=[string]$saved.Trabajo}
  if($saved.CSV){$script:csvFolder=[string]$saved.CSV}
  if($saved.Excel){$script:excelFolder=[string]$saved.Excel}
 }catch{}
}
$script:run = $null
$script:rows = $null
$script:files = @()
$script:big = @()
$script:small = @()
$script:parts = @()
$script:phase = 0
$script:begin = $null
$script:end = $null
$script:issues = New-Object 'System.Collections.Generic.List[string]'
$script:busy = $false

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Artikos | Banmedica" Width="1240" Height="830" MinWidth="490" MinHeight="520" ResizeMode="CanResize" WindowState="Normal" WindowStartupLocation="CenterScreen" Background="#FFFFFF" FontFamily="Segoe UI" UseLayoutRounding="True">
 <Window.Resources>
  <SolidColorBrush x:Key="Blue" Color="#0786F8"/>
  <SolidColorBrush x:Key="Green" Color="#00CB4D"/>
  <SolidColorBrush x:Key="Ink" Color="#100B59"/>
  <Style TargetType="Button" x:Key="PrimaryButton">
   <Setter Property="Foreground" Value="White"/><Setter Property="Background" Value="{StaticResource Blue}"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Padding" Value="18,11"/><Setter Property="FontSize" Value="13"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Cursor" Value="Hand"/><Setter Property="MinHeight" Value="43"/><Setter Property="Margin" Value="0,0,10,0"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Surface" Background="{TemplateBinding Background}" CornerRadius="10" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Surface" Property="Opacity" Value="0.82"/><Setter Property="Effect"><Setter.Value><DropShadowEffect Color="#4B8EA7" BlurRadius="12" ShadowDepth="2" Opacity="0.28"/></Setter.Value></Setter></Trigger><Trigger Property="IsPressed" Value="True"><Setter TargetName="Surface" Property="Opacity" Value="0.65"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter TargetName="Surface" Property="Background" Value="#B5C5CD"/><Setter Property="Cursor" Value="Arrow"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="TextBox"><Setter Property="Padding" Value="10,8"/><Setter Property="FontSize" Value="13"/><Setter Property="BorderBrush" Value="#D4E2E9"/><Setter Property="BorderThickness" Value="1"/><Setter Property="MinHeight" Value="37"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
  <Style TargetType="PasswordBox"><Setter Property="Padding" Value="10,8"/><Setter Property="FontSize" Value="13"/><Setter Property="BorderBrush" Value="#D4E2E9"/><Setter Property="BorderThickness" Value="1"/><Setter Property="MinHeight" Value="37"/></Style>
  <Style TargetType="DatePicker"><Setter Property="FontSize" Value="13"/><Setter Property="MinHeight" Value="37"/><Setter Property="BorderBrush" Value="#D4E2E9"/></Style>
  <Style TargetType="TextBlock" x:Key="FieldLabel"><Setter Property="FontSize" Value="12"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Foreground" Value="#5C7382"/><Setter Property="Margin" Value="0,0,0,6"/></Style>
  <Style x:Key="Card" TargetType="Border"><Setter Property="Background" Value="White"/><Setter Property="CornerRadius" Value="14"/><Setter Property="BorderBrush" Value="#E2EBEF"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="22"/><Setter Property="Margin" Value="0,0,0,18"/></Style>
 </Window.Resources>
 <Grid>
  <Grid.RowDefinitions><RowDefinition Height="92"/><RowDefinition Height="*"/><RowDefinition Height="32"/></Grid.RowDefinitions>
  <Border Background="White" BorderBrush="#E2EBEF" BorderThickness="0,0,0,1" Grid.Row="0">
   <Grid Margin="18,0"><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
     <Image x:Name="BrandImage" Width="195" Height="60" Stretch="Uniform" Visibility="Collapsed" Margin="0,0,22,0"/>
     <TextBlock x:Name="BrandText" Text="artikos" FontSize="32" FontWeight="Bold" Foreground="#0786F8" VerticalAlignment="Center" Margin="0,0,22,0"/>
     <Border Width="2" Background="#DCE9EF" Margin="0,13,20,13"/>
     <StackPanel x:Name="HeaderTitle" VerticalAlignment="Center"><TextBlock Text="Preparador de nóminas" FontSize="19" FontWeight="Bold" Foreground="#100B59"/><TextBlock Text="BANMÉDICA  ·  CONTROL OPERACIONAL" FontSize="11" Foreground="#6E8797" Margin="0,3,0,0"/></StackPanel>
    </StackPanel>
    <Border x:Name="HeaderStatus" Grid.Column="2" Background="#E7F7EE" CornerRadius="16" Padding="14,8" VerticalAlignment="Center"><StackPanel Orientation="Horizontal"><Ellipse Width="8" Height="8" Fill="#00CB4D" Margin="0,0,8,0"/><TextBlock Text="SOLO PREPARACIÓN LOCAL" FontSize="11" FontWeight="Bold" Foreground="#167447"/></StackPanel></Border>
   </Grid>
  </Border>
  <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"><StackPanel Margin="18,24,18,18">
   <TextBlock Text="Panel de preparación" FontSize="25" FontWeight="Bold" Foreground="#100B59"/>
   <TextBlock Text="Consulta, clasifica y prepara los archivos ZIP. La revisión final corresponde al operador disponible." FontSize="13" Foreground="#6C8290" Margin="0,4,0,22"/>
   <Border Style="{StaticResource Card}"><StackPanel>
    <Grid Margin="0,0,0,15"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Conexión y período" FontSize="17" FontWeight="Bold" Foreground="#100B59"/><Button x:Name="BtnWeek" Grid.Column="1" Style="{StaticResource PrimaryButton}" Background="#00CB4D" Content="Semana anterior" Padding="15,8" MinHeight="34" Margin="0"/></Grid>
    <WrapPanel x:Name="CredentialsPanel" ItemHeight="Auto" Margin="0,0,0,8">
     <StackPanel Width="215" Margin="0,0,12,12"><TextBlock Text="Servidor SQL" Style="{StaticResource FieldLabel}"/><TextBox x:Name="SqlServer" Text="192.168.70.27"/></StackPanel>
     <StackPanel Width="175" Margin="0,0,12,12"><TextBlock Text="Usuario SQL" Style="{StaticResource FieldLabel}"/><TextBox x:Name="SqlUser"/></StackPanel>
     <StackPanel Width="175" Margin="0,0,12,12"><TextBlock Text="Contraseña" Style="{StaticResource FieldLabel}"/><PasswordBox x:Name="SqlPassword"/></StackPanel>
     <StackPanel Width="175" Margin="0,0,12,12"><TextBlock Text="Fecha inicial" Style="{StaticResource FieldLabel}"/><DatePicker x:Name="DateFrom" SelectedDateFormat="Short"/></StackPanel>
     <StackPanel Width="175" Margin="0,0,12,12"><TextBlock Text="Fecha final" Style="{StaticResource FieldLabel}"/><DatePicker x:Name="DateTo" SelectedDateFormat="Short"/></StackPanel>
    </WrapPanel>
    <StackPanel Margin="0,12,0,0"><TextBlock Text="Distribución local · copie los ZIP manualmente" Style="{StaticResource FieldLabel}"/><TextBox x:Name="SourcePath" IsReadOnly="True"/><WrapPanel Margin="0,10,0,0"><Button x:Name="BtnBrowse" Style="{StaticResource PrimaryButton}" Content="Elegir carpeta"/><Button x:Name="BtnTest" Style="{StaticResource PrimaryButton}" Content="Probar SQL"/></WrapPanel></StackPanel>
    <Border Background="#F5FAFD" BorderBrush="#DCEAF1" BorderThickness="1" CornerRadius="10" Padding="16" Margin="0,18,0,0"><StackPanel>
     <TextBlock Text="Ubicaciones configurables · tú eliges dónde guardar cada resultado" FontWeight="SemiBold" Foreground="#100B59" Margin="0,0,0,12"/>
     <StackPanel Margin="0,0,0,12"><TextBlock Text="Carpeta de trabajo" Style="{StaticResource FieldLabel}"/><TextBox x:Name="WorkPath" IsReadOnly="True"/><Button x:Name="BtnWorkPath" Content="Examinar..." Style="{StaticResource PrimaryButton}" HorizontalAlignment="Left" Margin="0,8,0,0"/></StackPanel>
     <StackPanel Margin="0,0,0,12"><TextBlock Text="Guardar CSV en" Style="{StaticResource FieldLabel}"/><TextBox x:Name="CsvPath" IsReadOnly="True"/><Button x:Name="BtnCsvPath" Content="Examinar..." Style="{StaticResource PrimaryButton}" HorizontalAlignment="Left" Margin="0,8,0,0"/></StackPanel>
     <StackPanel Margin="0,0,0,12"><TextBlock Text="Guardar Excel en" Style="{StaticResource FieldLabel}"/><TextBox x:Name="ExcelPath" IsReadOnly="True"/><Button x:Name="BtnExcelPath" Content="Examinar..." Style="{StaticResource PrimaryButton}" HorizontalAlignment="Left" Margin="0,8,0,0"/></StackPanel>
     <TextBlock Text="NO SE GUARDAN CREDENCIALES NI DATOS SENSIBLES DE SQL." Foreground="#6C8290" FontSize="11" Margin="0,10,0,0"/>
    </StackPanel></Border>
   </StackPanel></Border>
   <WrapPanel Margin="0,0,0,18">
    <Border Width="220" Style="{StaticResource Card}" Margin="0,0,10,10"><StackPanel><TextBlock Text="REGISTROS SQL" Style="{StaticResource FieldLabel}"/><TextBlock x:Name="CountSql" Text="—" FontSize="30" FontWeight="Bold" Foreground="#0786F8"/></StackPanel></Border>
    <Border Width="220" Style="{StaticResource Card}" Margin="0,0,10,10"><StackPanel><TextBlock Text="DISTRIBUCIÓN" Style="{StaticResource FieldLabel}"/><TextBlock x:Name="CountAll" Text="—" FontSize="30" FontWeight="Bold" Foreground="#100B59"/></StackPanel></Border>
    <Border Width="220" Style="{StaticResource Card}" Margin="0,0,10,10"><StackPanel><TextBlock Text="GRANDES" Style="{StaticResource FieldLabel}"/><TextBlock x:Name="CountBig" Text="—" FontSize="30" FontWeight="Bold" Foreground="#00CB4D"/></StackPanel></Border>
    <Border Width="220" Style="{StaticResource Card}" Margin="0,0,10,10"><StackPanel><TextBlock Text="PEQUEÑOS" Style="{StaticResource FieldLabel}"/><TextBlock x:Name="CountSmall" Text="—" FontSize="30" FontWeight="Bold" Foreground="#0786F8"/></StackPanel></Border>
   </WrapPanel>
   <Border Style="{StaticResource Card}"><StackPanel>
    <TextBlock Text="Etapas de trabajo" FontSize="17" FontWeight="Bold" Foreground="#100B59" Margin="0,0,0,7"/>
    <TextBlock Text="Primero copie manualmente los ZIP del período en Distribución. Cada etapa solicita autorización." Foreground="#6C8290" Margin="0,0,0,17"/>
    <WrapPanel><Button x:Name="BtnQuery" Style="{StaticResource PrimaryButton}" Content="1  Consultar y generar Excel"/><Button x:Name="BtnCopy" Style="{StaticResource PrimaryButton}" Content="2  Clasificar carpeta local"/><Button x:Name="BtnCompress" Style="{StaticResource PrimaryButton}" Background="#00CB4D" Content="3  Preparar grandes"/><Button x:Name="BtnReview" Style="{StaticResource PrimaryButton}" Background="#100B59" Content="4  Validar y abrir carpeta"/></WrapPanel>
    <ProgressBar x:Name="Progress" Height="7" Margin="0,19,0,0" Minimum="0" Maximum="100" Value="0" Foreground="#00CB4D" Background="#E7EFF2"/>
    <TextBlock x:Name="Status" Text="Listo para comenzar." Foreground="#5C7382" Margin="0,10,0,0"/>
   </StackPanel></Border>
   <Border Style="{StaticResource Card}"><StackPanel><Grid Margin="0,0,0,12"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Bitácora" FontSize="17" FontWeight="Bold" Foreground="#100B59"/><Button x:Name="BtnFolder" Grid.Column="1" Content="Abrir trabajo" Style="{StaticResource PrimaryButton}" Background="#00CB4D" Padding="13,7" MinHeight="32" Margin="0"/></Grid><TextBox x:Name="LogBox" Height="145" Background="#F7FAFB" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/></StackPanel></Border>
  </StackPanel></ScrollViewer>
  <Border Grid.Row="2" Background="#EAF2F5"><TextBlock Text="ARTIKOS  ·  Preparación local de nóminas  ·  Uso exclusivamente interno por el área encargada" VerticalAlignment="Center" HorizontalAlignment="Center" Foreground="#67818F" FontSize="11"/></Border>
 </Grid>
</Window>
'@
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$script:window = [Windows.Markup.XamlReader]::Load($reader)
$names = 'BrandImage','BrandText','BtnWeek','SqlServer','SqlUser','SqlPassword','DateFrom','DateTo','SourcePath','BtnTest','BtnBrowse','CountSql','CountAll','CountBig','CountSmall','BtnQuery','BtnCopy','BtnCompress','BtnReview','Progress','Status','BtnFolder','LogBox','WorkPath','CsvPath','ExcelPath','BtnWorkPath','BtnCsvPath','BtnExcelPath','HeaderTitle','HeaderStatus'
foreach($n in $names){Set-Variable -Name $n -Value $script:window.FindName($n) -Scope Script}
$script:window.Add_SizeChanged({
 $narrow=$script:window.ActualWidth -lt 870
 $HeaderTitle.Visibility=if($narrow){'Collapsed'}else{'Visible'}
 $HeaderStatus.Visibility=if($narrow){'Collapsed'}else{'Visible'}
})
$SourcePath.Text=$script:inbox
$WorkPath.Text=$script:work
$CsvPath.Text=$script:csvFolder
$ExcelPath.Text=$script:excelFolder
$logo = Join-Path $script:root 'branding\artikos.png'
if(Test-Path -LiteralPath $logo){
 try{$bitmap=New-Object Windows.Media.Imaging.BitmapImage;$bitmap.BeginInit();$bitmap.CacheOption='OnLoad';$bitmap.UriSource=[Uri]::new($logo);$bitmap.EndInit();$BrandImage.Source=$bitmap;$BrandImage.Visibility='Visible';$BrandText.Visibility='Collapsed'}catch{}
}
function SetWeek { $s=[datetime]::Today; $s=$s.AddDays(-[int]$s.DayOfWeek); $DateTo.SelectedDate=$s; $DateFrom.SelectedDate=$s.AddDays(-7) }
SetWeek
function WriteAudit([string]$process,[string]$auditEvent,[hashtable]$details=@{},[string]$runOverride=$null){
 $targetRun=if($PSBoundParameters.ContainsKey('runOverride')){$runOverride}else{$script:run}
 $operationId=if($details.ContainsKey('operationId')){[string]$details.operationId}else{$null}
 $entry=[ordered]@{timestamp=(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffK');session=$script:sessionId;operationId=$operationId;process=$process;event=$auditEvent;user=($env:USERDOMAIN+'\'+$env:USERNAME);computer=$env:COMPUTERNAME;run=$targetRun;details=$details}
 $line=ConvertTo-Json -InputObject ([pscustomobject]$entry) -Compress -Depth 6
 $utf8=New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
 [IO.File]::AppendAllText($script:auditPath,$line+[Environment]::NewLine,$utf8)
 if($targetRun -and (Test-Path -LiteralPath $targetRun)){[IO.File]::AppendAllText((Join-Path $targetRun 'auditoria_procesos.jsonl'),$line+[Environment]::NewLine,$utf8)}
}
function StartAuditProcess([string]$name,[hashtable]$details=@{}){
 $token=@{Name=$name;Started=[datetime]::Now;Run=$script:run;Id=[guid]::NewGuid().ToString()}
 $details.operationId=$token.Id
 WriteAudit $name 'START' $details
 return $token
}
function CompleteAuditProcess($token,[hashtable]$details=@{}){
 $result=@{operationId=$token.Id;durationMs=[int]([datetime]::Now-$token.Started).TotalMilliseconds;outcome='success'}
 foreach($key in $details.Keys){$result[$key]=$details[$key]}
 if($result.outcome -eq 'success'){WriteAudit $token.Name 'COMPLETE' $result}else{WriteAudit $token.Name 'END' $result}
}
function FailAuditProcess($token,$err){
 if($token){WriteAudit $token.Name 'ERROR' @{operationId=$token.Id;durationMs=[int]([datetime]::Now-$token.Started).TotalMilliseconds;outcome='error';message=$err.Exception.Message} $token.Run}
}
function InvokeWinRAR([string]$executable,[string[]]$arguments,[string]$operation,[string]$fileName){
 $timer=[Diagnostics.Stopwatch]::StartNew();$code=$null;$operationId=[guid]::NewGuid().ToString();WriteAudit ('WINRAR_'+$operation) 'START' @{operationId=$operationId;file=$fileName;arguments=$arguments}
 try{
 & $executable @arguments | Out-Null
 $code=$LASTEXITCODE;$timer.Stop()
 if($code -ne 0){throw ('WinRAR '+$operation+' terminó con código '+$code+' para '+$fileName)}
 WriteAudit ('WINRAR_'+$operation) 'COMPLETE' @{operationId=$operationId;file=$fileName;exitCode=$code;durationMs=$timer.ElapsedMilliseconds}
 return $code
 }catch{
 $timer.Stop();WriteAudit ('WINRAR_'+$operation) 'ERROR' @{operationId=$operationId;file=$fileName;exitCode=$code;durationMs=$timer.ElapsedMilliseconds;message=$_.Exception.Message};throw
 }
}
function WriteLog([string]$message){
 $line='['+(Get-Date -Format 'HH:mm:ss')+'] '+$message
 $LogBox.AppendText($line+[Environment]::NewLine);$LogBox.ScrollToEnd();$Status.Text=$message
 WriteAudit 'APPLICATION' 'MESSAGE' @{message=$message}
 if($script:run){Add-Content -LiteralPath (Join-Path $script:run 'auditoria.log') -Value $line -Encoding UTF8}
 [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background)
}
function Ask([string]$message){return ([Windows.MessageBox]::Show($message+"`n`n¿Autorizar esta etapa?",'Autorización del operador','YesNo','Question') -eq 'Yes')}
function AlertError($err){$msg=$err.Exception.Message;WriteLog ('ERROR: '+$msg);[Windows.MessageBox]::Show($msg,'Error del proceso','OK','Error')|Out-Null}
function CheckDates {
 if(-not $DateFrom.SelectedDate -or -not $DateTo.SelectedDate){throw 'Seleccione ambos domingos.'}
 $script:begin=$DateFrom.SelectedDate.Date;$script:end=$DateTo.SelectedDate.Date
 if($script:begin.DayOfWeek -ne 'Sunday' -or $script:end.DayOfWeek -ne 'Sunday' -or $script:end -le $script:begin){throw 'El rango debe comenzar y terminar en domingo, y la fecha final debe ser posterior.'}
}
function GetSqlConnection {
 if([string]::IsNullOrWhiteSpace($SqlServer.Text) -or [string]::IsNullOrWhiteSpace($SqlUser.Text) -or [string]::IsNullOrWhiteSpace($SqlPassword.Password)){throw 'Ingrese servidor, usuario y contraseña SQL.'}
 # Utilizar las claves reconocidas por SqlClient; no utilizar DataSource como texto de clave.
 $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
 $builder['Data Source'] = $SqlServer.Text.Trim()
 $builder['Initial Catalog'] = 'SAF'
 $builder['User ID'] = $SqlUser.Text.Trim()
 $builder['Password'] = $SqlPassword.Password
 $builder['Integrated Security'] = $false
 $builder['Encrypt'] = $false
 $builder['TrustServerCertificate'] = $true
 $builder['Connect Timeout'] = 30
 return (New-Object System.Data.SqlClient.SqlConnection -ArgumentList $builder.ConnectionString)
}

function QuerySql {
 $audit=StartAuditProcess 'SQL_QUERY' @{server=$SqlServer.Text.Trim();database='SAF';from=$script:begin.ToString('o');to=$script:end.ToString('o')}
 $connection=$null
 try {
  $connection=GetSqlConnection
  $command=$connection.CreateCommand();$command.CommandTimeout=180
  $command.CommandText=@'
SELECT SPE.COD_EXTERNO AS [ID_EMPRESA], SPE.EMPRESA AS [EMPRESA], SSN.ID_NOMINA AS [NOMINA], SSN.TIPO AS [NOMBRE NOMINA], SSN.FH_GENERACION AS [FECHA GENERACION NOMINA], DSOL.FH_RESPUESTA AS [FH GENERACION PDF], CASE WHEN SSN.ID_ESTADO=101 THEN 'CONFIRMADA' END AS [ESTADO], SSN.NUM_DOCUMENTOS AS [DOCUMENTOS POR NOMINAS], CASE WHEN SSN.ARCHIVO_PDF_RESPALDO IS NULL THEN 'NO GENERADO' ELSE SSN.ARCHIVO_PDF_RESPALDO END AS [ARCHIVO_PDF_RESPALDO], DSOL.ID AS [ID SOLICITUD], CASE WHEN DSOL.ESTADO=1 THEN 'GENERADO' ELSE 'NO GENERADO' END AS [ESTADO DE SOLICITUD]
FROM SAF.DBO.SAF_NOMINA SSN
LEFT JOIN SAF..PAR_EMPRESAS SPE ON SSN.ID_EMPRESA=SPE.ID_EMPRESA
LEFT JOIN SAF..PAR_NOMINA_TIPO PNT ON SSN.COD_TIPO=PNT.COD_TIPO AND SSN.ID_EMPRESA=PNT.ID_EMPRESA
LEFT JOIN DTE_CONFIG..DTE_SOLICITUDES DSOL ON SSN.ID_NOMINA=DSOL.ID_DOCUMENTO
WHERE SSN.ID_EMPRESA IN (SELECT ID_EMPRESA FROM SAF..PAR_EMPRESAS WHERE RESPALDO_NOMINA_PDF='SI')
AND DSOL.FH_INGRESO >= @Inicio AND DSOL.FH_INGRESO < @Fin
AND SPE.COD_EXTERNO LIKE 'CL%' AND SSN.ID_ESTADO=101 AND DSOL.TIPO_SOLICITUD='SAF_NOMINA_PDFMASIVO' AND DSOL.ESTADO=1
ORDER BY DSOL.FH_INGRESO DESC
'@
  [void]$command.Parameters.Add('@Inicio',[Data.SqlDbType]::DateTime2);$command.Parameters['@Inicio'].Value=$script:begin
  [void]$command.Parameters.Add('@Fin',[Data.SqlDbType]::DateTime2);$command.Parameters['@Fin'].Value=$script:end
  $table=New-Object Data.DataTable
  $connection.Open();$reader=$command.ExecuteReader();$table.Load($reader);$reader.Close()
  CompleteAuditProcess $audit @{rows=$table.Rows.Count}
  return ,$table
 }catch{FailAuditProcess $audit $_;throw}
 finally{if($connection){$connection.Dispose()}}
}
function ExportReport([Data.DataTable]$table){
 $audit=StartAuditProcess 'REPORT_EXPORT' @{rows=$table.Rows.Count;csvDirectory=$script:csvFolder;excelDirectory=$script:excelFolder}
 try{
 [void](New-Item -ItemType Directory -Force -Path $script:csvFolder)
 $csv=Join-Path $script:csvFolder ('Reporte_Banmedica_'+$script:begin.ToString('yyyyMMdd')+'_'+$script:end.AddDays(-1).ToString('yyyyMMdd')+'_'+(Get-Date -Format 'HHmmss')+'.csv')
 $columns=@($table.Columns|ForEach-Object{$_.ColumnName})
 $records=@(foreach($r in $table.Rows){$obj=[ordered]@{};foreach($key in $columns){$obj[$key]=if($r[$key] -eq [DBNull]::Value){''}else{$r[$key]}};[pscustomobject]$obj})
 if($records.Count){$records|Export-Csv -LiteralPath $csv -Delimiter ';' -NoTypeInformation -Encoding UTF8}else{Set-Content -LiteralPath $csv -Value ($columns -join ';') -Encoding UTF8}
 WriteLog ('CSV generado: '+$csv)
 $template=Join-Path $script:root 'plantilla\Reporte_CSC_Banmedica.xlsx'
 if(!(Test-Path -LiteralPath $template)){WriteLog 'Plantilla Excel no encontrada; se conserva CSV.';CompleteAuditProcess $audit @{csv=$csv;rows=$table.Rows.Count;excel='template_missing'};return}
 $excelAudit=StartAuditProcess 'EXCEL_EXPORT' @{template=$template;rows=$table.Rows.Count}
 $excelStatus='failed'
 $excel=$null;$wb=$null;$sheet=$null
 try{
  $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
  $wb=$excel.Workbooks.Open($template);$sheet=$wb.Worksheets.Item(1)
  $last=[Math]::Max(5,$sheet.UsedRange.Row+$sheet.UsedRange.Rows.Count-1)
  $sheet.Range("C5:M$last").ClearContents() | Out-Null
  $sheet.Range('B2').Value2=$script:begin.ToOADate();$sheet.Range('B2').NumberFormat='dd/mm/yyyy'
  $sheet.Range('C2').Formula='=B2+6'
  $sheet.Range('D2').Formula=('=COUNT(E5:E'+[Math]::Max(5,$table.Rows.Count+4)+')')
  for($i=0;$i -lt $table.Rows.Count;$i++){
   for($j=0;$j -lt 11;$j++){
    $value=$table.Rows[$i][$j]
    if($value -ne [DBNull]::Value){$cell=$sheet.Cells.Item($i+5,$j+3);if($value -is [datetime]){$cell.Value2=$value.ToOADate();$cell.NumberFormat='dd/mm/yyyy hh:mm:ss'}else{$cell.Value2=([string]$value)}}
   }
   if(($i%75) -eq 0){$Progress.Value=5+[Math]::Floor(20*$i/[Math]::Max(1,$table.Rows.Count));[Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background)}
  }
  [void](New-Item -ItemType Directory -Force -Path $script:excelFolder)
  $out=Join-Path $script:excelFolder ('Reporte_CSC_Banmedica_'+$script:begin.ToString('yyyyMMdd')+'_'+$script:end.AddDays(-1).ToString('yyyyMMdd')+'.xlsx')
  $wb.SaveAs($out,51);$excelStatus='success';WriteLog ('Excel generado: '+$out);CompleteAuditProcess $excelAudit @{excel=$out;rows=$table.Rows.Count}
 }catch{FailAuditProcess $excelAudit $_;WriteLog ('No se pudo generar XLSX; CSV disponible. '+$_.Exception.Message)}
 finally{
  if($wb){$wb.Close($false)|Out-Null}
  if($excel){$excel.Quit()}
  foreach($obj in @($sheet,$wb,$excel)){if($null -ne $obj){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}catch{}}}
 }
 CompleteAuditProcess $audit @{csv=$csv;rows=$table.Rows.Count;excel=$excelStatus;excelPath=$(if($excelStatus -eq 'success'){$out}else{$null});outcome=$(if($excelStatus -eq 'success'){'success'}else{'partial'})}
}catch{FailAuditProcess $audit $_;throw}
}
function CopyChecked([string]$src,[string]$dst){
 $audit=StartAuditProcess 'FILE_COPY' @{source=$src;destination=$dst}
 try{
  if(Test-Path -LiteralPath $dst){throw ('Archivo ya existe, no se sobrescribe: '+$dst)}
  Copy-Item -LiteralPath $src -Destination $dst -ErrorAction Stop
  $a=Get-Item -LiteralPath $src;$b=Get-Item -LiteralPath $dst
  if($a.Length -ne $b.Length){throw ('Tamaño no coincide: '+$src)}
  $sourceHash=(Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
  $destinationHash=(Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash
  if($sourceHash -ne $destinationHash){throw ('SHA256 no coincide: '+$src)}
  CompleteAuditProcess $audit @{bytes=$b.Length;sha256=$sourceHash}
 }catch{FailAuditProcess $audit $_;throw}
}
function FindWinRAR {
 $candidates=@((Join-Path $env:ProgramFiles 'WinRAR\WinRAR.exe'))
 if(${env:ProgramFiles(x86)}){$candidates+=(Join-Path ${env:ProgramFiles(x86)} 'WinRAR\WinRAR.exe')}
 foreach($p in $candidates){if(Test-Path -LiteralPath $p){return $p}}
 $cmd=Get-Command WinRAR.exe -ErrorAction SilentlyContinue;if($cmd){return $cmd.Source}
 throw 'No se encontró WinRAR.exe. Verifique su instalación.'
}
function SetPhase([int]$n){$script:phase=$n;$BtnQuery.IsEnabled=($n -eq 0);$BtnCopy.IsEnabled=($n -eq 1);$BtnCompress.IsEnabled=($n -eq 2);$BtnReview.IsEnabled=($n -ge 2)}
# Inicio de sesión: autenticación real con SQL Server. No almacena contraseñas.
[xml]$loginXaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Artikos · Acceso seguro" Width="510" Height="560" MinWidth="380" MinHeight="460" WindowStartupLocation="CenterScreen" ResizeMode="CanResize" Background="#F5F9FC" FontFamily="Segoe UI" UseLayoutRounding="True">
 <Window.Resources>
  <Style TargetType="TextBox"><Setter Property="Padding" Value="12,10"/><Setter Property="MinHeight" Value="43"/><Setter Property="FontSize" Value="14"/><Setter Property="BorderBrush" Value="#D3E3EB"/></Style>
  <Style TargetType="PasswordBox"><Setter Property="Padding" Value="12,10"/><Setter Property="MinHeight" Value="43"/><Setter Property="FontSize" Value="14"/><Setter Property="BorderBrush" Value="#D3E3EB"/></Style>
 </Window.Resources>
 <Grid><ScrollViewer VerticalScrollBarVisibility="Auto"><Border Margin="22" Background="White" BorderBrush="#E0EAF0" BorderThickness="1" CornerRadius="18" Padding="24"><StackPanel>
  <Image Name="LoginLogo" Height="72" Stretch="Uniform" HorizontalAlignment="Left"/>
  <TextBlock Text="Bienvenido" FontSize="27" FontWeight="Bold" Foreground="#100B59" Margin="0,18,0,4"/>
  <TextBlock Text="Centro de preparación de nóminas · Banmédica" FontSize="13" Foreground="#5D7485" TextWrapping="Wrap" Margin="0,0,0,20"/>
  <TextBlock Text="SERVIDOR SQL" Foreground="#526B7B" FontWeight="SemiBold" Margin="0,0,0,5"/><TextBox Name="LoginServer" Text="192.168.70.27"/>
  <TextBlock Text="USUARIO" Foreground="#526B7B" FontWeight="SemiBold" Margin="0,14,0,5"/><TextBox Name="LoginUser"/>
  <TextBlock Text="CONTRASEÑA" Foreground="#526B7B" FontWeight="SemiBold" Margin="0,14,0,5"/><PasswordBox Name="LoginPassword"/>
  <TextBlock Name="LoginStatus" Foreground="#B32222" TextWrapping="Wrap" Margin="0,12,0,0"/>
  <Button Name="LoginEnter" Content="Iniciar sesión" Height="48" Margin="0,20,0,8" Background="#0786F8" Foreground="White" FontWeight="Bold" FontSize="14" BorderThickness="0" Cursor="Hand"/>
  <Button Name="LoginCancel" Content="Salir" Height="39" Background="#EEF5F9" Foreground="#100B59" BorderThickness="0" Cursor="Hand"/>
  <TextBlock Text="Acceso validado contra SQL Server. No se almacena la contraseña." TextWrapping="Wrap" TextAlignment="Center" FontSize="11" Foreground="#778D9A" Margin="0,15,0,0"/>
 </StackPanel></Border></ScrollViewer></Grid>
</Window>
'@
$loginWindow=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader($loginXaml)))
$loginLogo=$loginWindow.FindName('LoginLogo')
if(Test-Path $logo){$loginLogo.Source=$BrandImage.Source}
$loginServer=$loginWindow.FindName('LoginServer');$loginUser=$loginWindow.FindName('LoginUser');$loginPassword=$loginWindow.FindName('LoginPassword');$loginStatus=$loginWindow.FindName('LoginStatus')
$loginWindow.FindName('LoginCancel').Add_Click({$loginWindow.DialogResult=$false;$loginWindow.Close()})
$loginWindow.FindName('LoginEnter').Add_Click({
 $audit=StartAuditProcess 'SQL_LOGIN' @{server=$loginServer.Text.Trim();database='SAF';user=$loginUser.Text.Trim()}
 try {
  $SqlServer.Text=$loginServer.Text;$SqlUser.Text=$loginUser.Text;$SqlPassword.Password=$loginPassword.Password
  $c=GetSqlConnection
  try{$c.Open()}finally{$c.Dispose()}
  CompleteAuditProcess $audit
  $loginWindow.DialogResult=$true;$loginWindow.Close()
 }catch{FailAuditProcess $audit $_;$loginStatus.Text='No fue posible iniciar sesión: '+$_.Exception.Message}
})
if($loginWindow.ShowDialog() -ne $true){WriteAudit 'APPLICATION' 'LOGIN_CANCELLED' @{};return}
$SqlUser.IsReadOnly=$true;$SqlServer.IsReadOnly=$true;$SqlPassword.IsEnabled=$false

$BtnBrowse.Add_Click({
 $audit=StartAuditProcess 'SELECT_SOURCE_FOLDER' @{}
 try{
  Add-Type -AssemblyName System.Windows.Forms
  $dlg=New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description='Seleccione la carpeta local con los ZIP originales'
  $dlg.SelectedPath=$SourcePath.Text
  if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){$SourcePath.Text=$dlg.SelectedPath;CompleteAuditProcess $audit @{path=$dlg.SelectedPath}}else{CompleteAuditProcess $audit @{outcome='cancelled'}}
 }catch{FailAuditProcess $audit $_;AlertError $_}
})
function SavePreferences {
 $audit=StartAuditProcess 'SAVE_PREFERENCES' @{settingsFile=$script:settingsFile}
 try{
  $prefs=[ordered]@{Distribucion=$SourcePath.Text;Trabajo=$WorkPath.Text;CSV=$CsvPath.Text;Excel=$ExcelPath.Text}
  $preferencesJson=ConvertTo-Json -InputObject $prefs -Compress
  $preferencesHash=[Security.Cryptography.SHA256]::Create()
  try{$preferencesChecksum=([BitConverter]::ToString($preferencesHash.ComputeHash([Text.Encoding]::UTF8.GetBytes($preferencesJson)))).Replace('-','')}finally{$preferencesHash.Dispose()}
  $folder=Split-Path -Parent $script:settingsFile
  [void](New-Item -ItemType Directory -Force -Path $folder)
  $prefs | ConvertTo-Json | Set-Content -LiteralPath $script:settingsFile -Encoding UTF8 -ErrorAction Stop
  CompleteAuditProcess $audit @{settingsFile=$script:settingsFile;preferencesSha256=$preferencesChecksum}
 }catch{FailAuditProcess $audit $_;throw}
}
function ChoosePath([string]$current,[string]$title){
 $audit=StartAuditProcess 'SELECT_OUTPUT_FOLDER' @{purpose=$title;currentPath=$current}
 try{
  Add-Type -AssemblyName System.Windows.Forms
  $dlg=New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description=$title
  if(Test-Path -LiteralPath $current){$dlg.SelectedPath=$current}
  if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){CompleteAuditProcess $audit @{path=$dlg.SelectedPath};return $dlg.SelectedPath}
  CompleteAuditProcess $audit @{outcome='cancelled'}
  return $null
 }catch{FailAuditProcess $audit $_;throw}
}
$BtnWorkPath.Add_Click({$p=ChoosePath $WorkPath.Text 'Seleccione dónde guardar las ejecuciones y archivos preparados';if($p){$WorkPath.Text=$p;$script:work=$p;SavePreferences}})
$BtnCsvPath.Add_Click({$p=ChoosePath $CsvPath.Text 'Seleccione dónde guardar los archivos CSV';if($p){$CsvPath.Text=$p;$script:csvFolder=$p;SavePreferences}})
$BtnExcelPath.Add_Click({$p=ChoosePath $ExcelPath.Text 'Seleccione dónde guardar los archivos Excel';if($p){$ExcelPath.Text=$p;$script:excelFolder=$p;SavePreferences}})
$SourcePath.Add_TextChanged({if($SourcePath.Text){SavePreferences}})
$BtnWeek.Add_Click({$audit=StartAuditProcess 'SET_DEFAULT_DATE_RANGE' @{};try{SetWeek;CompleteAuditProcess $audit @{from=$DateFrom.SelectedDate.ToString('o');to=$DateTo.SelectedDate.ToString('o')}}catch{FailAuditProcess $audit $_;AlertError $_}})
$BtnTest.Add_Click({$audit=StartAuditProcess 'SQL_CONNECTION_TEST' @{server=$SqlServer.Text.Trim();database='SAF'};$c=$null;try{$c=GetSqlConnection;$c.Open();CompleteAuditProcess $audit;WriteLog 'Conexión SQL comprobada.';[Windows.MessageBox]::Show('Conexión correcta a SAF.','SQL Server','OK','Information')|Out-Null}catch{FailAuditProcess $audit $_;AlertError $_}finally{if($c){$c.Dispose()}}})
$BtnQuery.Add_Click({$audit=StartAuditProcess 'QUERY_AND_REPORT_STAGE' @{};try{
 CheckDates
 if(!(Ask ('Consultar SQL desde '+$script:begin.ToString('dd/MM/yyyy')+' hasta '+$script:end.AddDays(-1).ToString('dd/MM/yyyy')+' y generar el reporte local.'))){CompleteAuditProcess $audit @{outcome='cancelled';reason='operator_denied'};return}
 $base=$script:begin.ToString('yyyyMMdd')+'_'+$script:end.AddDays(-1).ToString('yyyyMMdd')
 $script:run=Join-Path $script:work ($base+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))
 [void](New-Item -ItemType Directory -Force -Path $script:run)
 $Progress.Value=2;WriteLog 'Consultando SQL Server...'
 $script:rows=QuerySql
 $CountSql.Text=[string]$script:rows.Rows.Count
 ExportReport $script:rows;$Progress.Value=25;WriteLog ('Consulta completada: '+$script:rows.Rows.Count+' registros.');SetPhase 1;CompleteAuditProcess $audit @{rows=$script:rows.Rows.Count;run=$script:run}
}catch{FailAuditProcess $audit $_;AlertError $_}})
$BtnCopy.Add_Click({$audit=StartAuditProcess 'ZIP_CLASSIFY_AND_COPY' @{sourceDirectory=$SourcePath.Text};try{
 if(!(Test-Path -LiteralPath $SourcePath.Text)){throw 'No se encuentra la carpeta local Distribución.'}
 $all=@(Get-ChildItem -LiteralPath $SourcePath.Text -File -Filter '*.zip' | Sort-Object -Property @{Expression='Length';Descending=$true},@{Expression='Name';Descending=$false})
 if($all.Count -eq 0){throw 'No se encontraron ZIP en la carpeta local Distribución.'}
 $expected=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
 foreach($r in $script:rows.Rows){
  $raw=[string]$r['ARCHIVO_PDF_RESPALDO']
  if($raw -and $raw -ne 'NO GENERADO'){
   foreach($match in [regex]::Matches($raw,'(?i)CL[A-Z0-9_\-]+\.zip')){[void]$expected.Add($match.Value)}
  }
 }
 if($expected.Count -eq 0){throw 'SQL no entregó nombres ZIP reconocibles en ARCHIVO_PDF_RESPALDO. Revise el reporte; no se clasificará sin validación.'}
 $found=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
 foreach($f in $all){[void]$found.Add($f.Name)}
 $missing=@($expected|Where-Object{-not $found.Contains($_)})
 $extra=@($found|Where-Object{-not $expected.Contains($_)})
 WriteLog ('Validación previa: SQL '+$expected.Count+' ZIP únicos, locales '+$found.Count+', faltantes '+$missing.Count+', adicionales '+$extra.Count)
 $script:issues.Clear()
 foreach($name in $missing){$script:issues.Add('FALTANTE;'+$name)}
 foreach($name in $extra){$script:issues.Add('ADICIONAL;'+$name)}
 $script:issues | Set-Content -LiteralPath (Join-Path $script:run 'Diferencias_SQL.txt') -Encoding UTF8
 if($missing.Count -or $extra.Count){
  CompleteAuditProcess $audit @{outcome='blocked';reason='sql_file_mismatch';missing=$missing.Count;extra=$extra.Count}
  [Windows.MessageBox]::Show(('Se detectaron diferencias: '+$missing.Count+' faltantes y '+$extra.Count+' adicionales. Revise Diferencias_SQL.txt. No se clasificarán archivos hasta corregirlas.'),'Validación SQL vs ZIP','OK','Warning')|Out-Null
  return
 }
 if(!(Ask ('Se encontraron '+$all.Count+' ZIP locales. Se conservarán los originales, se copiarán los 40 más pesados a Grande y los restantes a Pequeño.'))){CompleteAuditProcess $audit @{outcome='cancelled';files=$all.Count};return}
 foreach($folder in @('Distribucion','Grande','Pequeno','Preparados','Extraidos')){[void](New-Item -ItemType Directory -Force -Path (Join-Path $script:run $folder))}
 $rar=FindWinRAR
 $script:files=$all;$script:big=@();$script:small=@()
 for($i=0;$i -lt $all.Count;$i++){
  $file=$all[$i]
  $rarCode=InvokeWinRAR $rar @('t','-ibck','-y',$file.FullName) 'SOURCE_TEST' $file.Name
  if($rarCode -ne 0){throw ('ZIP de origen dañado: '+$file.Name+' (WinRAR '+$rarCode+')')}
  $local=Join-Path (Join-Path $script:run 'Distribucion') $file.Name
  CopyChecked $file.FullName $local
  $group=if($i -lt 40){'Grande'}else{'Pequeno'}
  $target=Join-Path (Join-Path $script:run $group) $file.Name
  CopyChecked $local $target
  if($i -lt 40){$script:big+=Get-Item -LiteralPath $target}else{$script:small+=Get-Item -LiteralPath $target}
  $Progress.Value=25+[Math]::Floor(35*($i+1)/$all.Count)
  if((($i+1)%20) -eq 0 -or $i -eq ($all.Count-1)){WriteLog ('Archivos copiados: '+($i+1)+' / '+$all.Count)}
 }
 $CountAll.Text=[string]$all.Count;$CountBig.Text=[string]$script:big.Count;$CountSmall.Text=[string]$script:small.Count
 WriteLog ('Clasificación terminada. Diferencias SQL: '+$script:issues.Count);SetPhase 2;CompleteAuditProcess $audit @{files=$all.Count;large=$script:big.Count;small=$script:small.Count;issues=$script:issues.Count}
}catch{FailAuditProcess $audit $_;AlertError $_}})
$BtnCompress.Add_Click({$audit=StartAuditProcess 'ZIP_REPACK_AND_SPLIT' @{fileCount=$script:big.Count;volumeLimitBytes=2097152};try{
 $rar=FindWinRAR
 if(!(Ask ('Extraer y recomprimir '+$script:big.Count+' ZIP grandes en volúmenes de 2 MiB. No se modifican los originales.'))){CompleteAuditProcess $audit @{outcome='cancelled';files=$script:big.Count};return}
 $dest=Join-Path $script:run 'Preparados\Grande';[void](New-Item -ItemType Directory -Force -Path $dest)
 for($i=0;$i -lt $script:big.Count;$i++){
  $f=$script:big[$i];$stem=[IO.Path]::GetFileNameWithoutExtension($f.Name)
  $folder=Join-Path (Join-Path $script:run 'Extraidos') $stem
  [void](New-Item -ItemType Directory -Force -Path $folder)
  $rarCode=InvokeWinRAR $rar @('x','-ibck','-y',$f.FullName,($folder+[IO.Path]::DirectorySeparatorChar)) 'EXTRACT' $f.Name
  if($rarCode -ne 0){throw ('Error al extraer: '+$f.Name+' (WinRAR '+$rarCode+')')}
  $output=Join-Path $dest $f.Name
  Push-Location $folder
  try{$rarCode=InvokeWinRAR $rar @('a','-afzip','-m5','-v1900k','-ibck','-y',$output,'*') 'COMPRESS' $f.Name;if($rarCode -ne 0){throw ('Error al comprimir: '+$f.Name+' (WinRAR '+$rarCode+')')}}finally{Pop-Location}
  $parts=@(Get-ChildItem -LiteralPath $dest -File | Where-Object{$_.Name -eq ($stem+'.zip') -or $_.Name -match ('^'+[regex]::Escape($stem)+'\.z\d+$')})
  if(!$parts.Count){throw ('No se generaron volúmenes: '+$stem)}
  if(!($parts | Where-Object {$_.Extension -eq '.zip'})){throw ('Falta ZIP final del conjunto: '+$stem)}
  foreach($part in $parts){if($part.Length -gt 2097152){throw ('Volumen supera 2 MiB: '+$part.Name)}}
  $rarCode=InvokeWinRAR $rar @('t','-ibck','-y',$output) 'OUTPUT_TEST' $f.Name
  if($rarCode -ne 0){throw ('Los volúmenes no superaron prueba de integridad: '+$f.Name)}
  $Progress.Value=60+[Math]::Floor(35*($i+1)/[Math]::Max(1,$script:big.Count))
  WriteLog ('Preparado '+($i+1)+'/'+$script:big.Count+': '+$stem+' ('+$parts.Count+' partes)')
 }
 $script:parts=@(Get-ChildItem -LiteralPath $dest -File)
 WriteLog ('Compresión terminada: '+$script:parts.Count+' volúmenes.');$Progress.Value=95;SetPhase 3;CompleteAuditProcess $audit @{files=$script:big.Count;volumes=$script:parts.Count}
}catch{FailAuditProcess $audit $_;AlertError $_}})
$BtnReview.Add_Click({$audit=StartAuditProcess 'INTEGRITY_REVIEW' @{run=$script:run};try{
 $original=@(Get-ChildItem -LiteralPath (Join-Path $script:run 'Distribucion') -File -Filter '*.zip')
 $big=@(Get-ChildItem -LiteralPath (Join-Path $script:run 'Grande') -File -Filter '*.zip')
 $small=@(Get-ChildItem -LiteralPath (Join-Path $script:run 'Pequeno') -File -Filter '*.zip')
 $problems=New-Object 'System.Collections.Generic.List[string]'
 if($original.Count -ne ($big.Count+$small.Count)){$problems.Add('Cantidad de ZIP no coincide.')}
 foreach($f in @($big)+@($small)){$src=Join-Path (Join-Path $script:run 'Distribucion') $f.Name;if(!(Test-Path -LiteralPath $src)){$problems.Add('Falta original: '+$f.Name)}elseif((Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash){$problems.Add('SHA256 diferente: '+$f.Name)}}
 if($script:phase -ge 3){foreach($p in $script:parts){if($p.Length -gt 2097152){$problems.Add('Volumen superior a 2 MiB: '+$p.Name)}}}
 $report=@('REVISIÓN OPERACIONAL - BANMÉDICA','Periodo: '+$script:begin.ToString('dd/MM/yyyy')+' a '+$script:end.AddDays(-1).ToString('dd/MM/yyyy'),'','SQL: '+$script:rows.Rows.Count,'Distribución: '+$original.Count,'Grandes: '+$big.Count,'Pequeños: '+$small.Count,'Volúmenes preparados: '+$script:parts.Count,'Diferencias SQL: '+$script:issues.Count,'Problemas de integridad: '+$problems.Count,'','No se realizaron envíos.','', 'DIFERENCIAS SQL')+@($script:issues)+@('','PROBLEMAS')+@($problems)
 $report|Set-Content -LiteralPath (Join-Path $script:run 'Revision_Operador.txt') -Encoding UTF8
 $Progress.Value=100;WriteLog ('Revisión lista. Incidencias: '+($script:issues.Count+$problems.Count));CompleteAuditProcess $audit @{sqlIssues=$script:issues.Count;integrityProblems=$problems.Count;resultFile=(Join-Path $script:run 'Revision_Operador.txt')};$openAudit=StartAuditProcess 'OPEN_RUN_FOLDER' @{path=$script:run};try{Start-Process explorer.exe -ArgumentList ('"'+$script:run+'"');CompleteAuditProcess $openAudit}catch{FailAuditProcess $openAudit $_;throw}
}catch{FailAuditProcess $audit $_;AlertError $_}})
$BtnFolder.Add_Click({$path=if($script:run){$script:run}else{$script:work};$audit=StartAuditProcess 'OPEN_WORK_FOLDER' @{path=$path};try{if(!(Test-Path -LiteralPath $path)){[void](New-Item -ItemType Directory -Force -Path $path)};Start-Process explorer.exe -ArgumentList ('"'+$path+'"');CompleteAuditProcess $audit}catch{FailAuditProcess $audit $_;AlertError $_}})
SetPhase 0
WriteAudit 'APPLICATION' 'SESSION_START' @{version='1.0'}
WriteLog 'Aplicación iniciada. Preparación local exclusivamente.'
try{[void]$script:window.ShowDialog()}finally{WriteAudit 'APPLICATION' 'SESSION_END' @{}}
