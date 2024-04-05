$excelPath = "C:\Users\dsalas\OneDrive - AUTORENTAS DEL PACIFICO LTDA\Escritorio\David\AvisoCaidaEnlace\Operaciones_TI-Hostname_Sucursales_y_Telefono.xlsx" # Ruta hacia el excel
$destinatarios = @("dsalas@autorentas.cl")
$remitente = "Contingencia Operacional <contingenciaoperacional@mitta.cl>" # Quien es el que envia el correo
$smtpServer = "......." # IP del Relay SMTP, es donde se envia el correo al server para luego este lo envie a Office365
$smtpPort = 25 #Puerto al cual se envia los correo
$imagPath = "C:\Users\dsalas\OneDrive - AUTORENTAS DEL PACIFICO LTDA\Escritorio\David\AvisoCaidaEnlace\Encabezado.png" # Ruta a la imagen a utilizar, debe estar en la misma carpeta. En este caso es encabezado de p�gina
$imagPath2 = "C:\Users\dsalas\OneDrive - AUTORENTAS DEL PACIFICO LTDA\Escritorio\David\AvisoCaidaEnlace\Footer.png" # Ruta a la imagen a utilizar, debe estar en la misma carpeta. En este caso es pie de p�gina

# Leer el archivo Excel y obtener los datos
try {
    $data = Import-Excel -Path $excelPath -WorksheetName "Prueba"
} catch {
    Write-Host "Error al leer el archivo Excel: $_.Exception.Message"
    exit
}

# Bucle Infinito
while ($true) {
    # Iterar a trav�s de cada fila en los datos
    foreach ($row in $data) {
        $hostname = $row.Hostname
        $sucursal = $row.Sucursal
        $telefono = $row.Telefono
        $correoSucursal = $row.CorreoSucursal

        try {
            $pingResult = Test-Connection -ComputerName $hostname -Count 2 -BufferSize 1024 -ErrorAction Stop

            # Verificaci�n de ping al Hostname(IP) de la sucursal
            if ($null -ne $pingResult.ResponseTime) {
                Write-Host "$sucursal ($hostname) $telefono UP"
            
                # Actualizar estado a "UP" en la fila
                $row | Add-Member -MemberType NoteProperty -Name Estado -Value "UP" -Force
            
                # Restablecer el contador de errores
                $row.ContadorErrores = 0
            
                # Verificaci�n si el EstadoAnterior est� DOWN
                if ($row.EstadoAnterior -eq "DOWN") {
                    Write-Host "�Cambio de estado! Enviando correo electr�nico..."
            
                    # Resto del c�digo para enviar el correo electr�nico...
            
                    # Convertir las im�genes a Base64
                    $base64Imagen1 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($imagPath))
                    $base64Imagen2 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($imagPath2))
            
                    # Configurar opciones del correo electr�nico
                    $correoOpciones = @{
                        To         = 'contingenciaoperacional@mitta.cl'
                        From       = $remitente
                        Subject    = "Sucursal $sucursal Operativo " + [char]::ConvertFromUtf32(0x26A1) + [char]::ConvertFromUtf32(0x1F50C)
                        Body       = ""
                        BodyAsHtml = $true
                        Encoding    = [System.Text.Encoding]::UTF8  # Agregar esta l�nea para especificar la codificaci�n
                        SmtpServer = $smtpServer
                        Port       = $smtpPort
                    }
            
                    # Embeber las URLs de las im�genes en el cuerpo del correo
                    $correoOpciones.Body = @"
                    <!doctype html>
                    <html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
                    
                    <head>
                      <title>
                      </title>
                      <!--[if !mso]><!-->
                      <meta http-equiv="X-UA-Compatible" content="IE=edge">
                      <!--<![endif]-->
                      <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
                      <meta name="viewport" content="width=device-width, initial-scale=1">
                      <style type="text/css">
                        #outlook a {
                          padding: 0;
                        }
                    
                        body {
                          margin: 0;
                          padding: 0;
                          -webkit-text-size-adjust: 100%;
                          -ms-text-size-adjust: 100%;
                        }
                    
                        table,
                        td {
                          border-collapse: collapse;
                          mso-table-lspace: 0pt;
                          mso-table-rspace: 0pt;
                        }
                    
                        img {
                          border: 0;
                          height: auto;
                          line-height: 100%;
                          outline: none;
                          text-decoration: none;
                          -ms-interpolation-mode: bicubic;
                        }
                    
                        p {
                          display: block;
                          margin: 13px 0;
                        }
                      </style>
                      <!--[if mso]>
                            <noscript>
                            <xml>
                            <o:OfficeDocumentSettings>
                              <o:AllowPNG/>
                              <o:PixelsPerInch>96</o:PixelsPerInch>
                            </o:OfficeDocumentSettings>
                            </xml>
                            </noscript>
                            <![endif]-->
                      <!--[if lte mso 11]>
                            <style type="text/css">
                              .mj-outlook-group-fix { width:100% !important; }
                            </style>
                            <![endif]-->
                      <!--[if !mso]><!-->
                      <link href="https://fonts.googleapis.com/css?family=Droid+Sans:300,400,500,700" rel="stylesheet" type="text/css">
                      <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700" rel="stylesheet" type="text/css">
                      <link href="https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700" rel="stylesheet" type="text/css">
                      <style type="text/css">
                        @import url(https://fonts.googleapis.com/css?family=Droid+Sans:300,400,500,700);
                        @import url(https://fonts.googleapis.com/css?family=Roboto:300,400,500,700);
                        @import url(https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700);
                      </style>
                      <!--<![endif]-->
                      <style type="text/css">
                        @media only screen and (min-width:600px) {
                          .mj-column-per-100 {
                            width: 100% !important;
                            max-width: 100%;
                          }
                        }
                      </style>
                      <style media="screen and (min-width:600px)">
                        .moz-text-html .mj-column-per-100 {
                          width: 100% !important;
                          max-width: 100%;
                        }
                      </style>
                      <style type="text/css">
                        @media only screen and (max-width:600px) {
                          table.mj-full-width-mobile {
                            width: 100% !important;
                          }
                    
                          td.mj-full-width-mobile {
                            width: auto !important;
                          }
                        }
                      </style>
                    </head>
                    
                    <body style="word-spacing:normal;background-color:#F7F7F7;">
                      <div style="background-color:#F7F7F7;">
                        <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
                        <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
                          <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
                            <tbody>
                              <tr>
                                <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
                                  <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
                                  <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                                      <tbody>
                                        <tr>
                                          <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                                            <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                                              <tbody>
                                                <tr>
                                                  <td align="center" style="font-size:0px;padding:0px 0px 0px 0px;word-break:break-word;">
                                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;border-spacing:0px;">
                                                      <tbody>
                                                        <tr>
                                                          <td style="width:800px;">
                                                            <img alt="imagen" height="auto" src="data:image/png;base64,$base64Imagen1" style="border:0;display:block;outline:none;text-decoration:none;height:auto;width:100%;font-size:13px;" width="800" />
                                                          </td>
                                                        </tr>
                                                      </tbody>
                                                    </table>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>
                                          </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </div>
                                  <!--[if mso | IE]></td></tr></table><![endif]-->
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <!--[if mso | IE]></td></tr></table><![endif]-->
                        <!-- BODY -->
                        <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
                        <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
                          <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
                            <tbody>
                              <tr>
                                <td style="direction:ltr;font-size:0px;padding:20px 0;text-align:center;">
                                  <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
                                  <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%">
                                      <tbody>
                                        <tr>
                                          <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                                            <div style="font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen', 'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans','Helvetica Neue', sans-serif;font-size:14px;font-weight:400;line-height:18px;text-align:left;color:#000000;">
                                              <p>Estimados colaboradores:</p>
                                              <p>Informamos que la <strong>sucursal $sucursal</strong> ha vuelto a estar en linea.</p>
                                              <p>Gracias por su paciencia.</p>
                                              <p>Atentamente,</p>
                                            </div>
                                          </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </div>
                                  <!--[if mso | IE]></td></tr></table><![endif]-->
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <!--[if mso | IE]></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
                        <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
                          <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
                            <tbody>
                              <tr>
                                <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
                                  <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
                                  <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                                      <tbody>
                                        <tr>
                                          <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                                            <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                                              <tbody>
                                              </tbody>
                                            </table>
                                          </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </div>
                                  <!--[if mso | IE]></td></tr></table><![endif]-->
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <!--[if mso | IE]></td></tr></table><![endif]-->
                        <!-- END BODY -->
                        <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
                        <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
                          <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
                            <tbody>
                              <tr>
                                <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
                                  <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
                                  <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                                      <tbody>
                                        <tr>
                                          <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                                            <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                                              <tbody>
                                                <tr>
                                                  <td align="center" style="font-size:0px;padding:0px 0px 0px 0px;word-break:break-word;">
                                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;border-spacing:0px;">
                                                      <tbody>
                                                        <tr>
                                                          <td style="width:800px;">
                                                            <img alt="imagen" height="auto" src="data:image/png;base64,$base64Imagen1" style="border:0;display:block;outline:none;text-decoration:none;height:auto;width:100%;font-size:13px;" width="800" />
                                                          </td>
                                                        </tr>
                                                      </tbody>
                                                    </table>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>
                                          </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </div>
                                  <!--[if mso | IE]></td></tr></table><![endif]-->
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <!--[if mso | IE]></td></tr></table><![endif]-->
                      </div>
                    </body>
                    
                    </html>
"@
                  # Enviar el correo electr�nico
                    Send-MailMessage @correoOpciones -Bcc ($destinatarios+$correoSucursal) -ErrorAction Stop
                    Write-Host "Correo electr�nico enviado a $($destinatarioArray -join ',') por cambio a UP."
                    $row | Add-Member -MemberType NoteProperty -Name EstadoAnterior -Value "UP" -Force
                }
            }
            
        } catch {
            #Escribe en consola la sucursal que DOWN
            Write-Host "Error al realizar el ping a $sucursal ($hostname): $_.Exception.Message"
            Write-Host "$sucursal ($hostname) $telefono DOWN"

            # Actualizar estado a "DOWN" en la fila
            $row | Add-Member -MemberType NoteProperty -Name Estado -Value "DOWN" -Force

            # Incrementar el contador de errores
            $row | Add-Member -MemberType NoteProperty -Name ContadorErrores -Value ($row.ContadorErrores + 1) -Force

            # Verificar si se super� el l�mite de errores (5 en este caso)
            if ($row.ContadorErrores -eq 1) {
                Write-Host "Enviar correo electr�nico por l�mite de errores alcanzado."
            
                # Configurar opciones del correo electr�nico
                $correoOpciones = @{
                    To         = 'contingenciaoperacional@mitta.cl'
                    From       = $remitente
                    Subject    = "Sucursal $sucursal Sin Conexion " + [char]::ConvertFromUtf32(0x26A1) + [char]::ConvertFromUtf32(0x1F50C)
                    BodyAsHtml = $true
                    SmtpServer = $smtpServer
                    Port       = $smtpPort
                    Encoding   = [System.Text.Encoding]::UTF8  # Agregar esta línea para especificar la codificación
                }
                
            
                
                # Convertir las im�genes a Base64
                $base64Imagen1 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($imagPath))
                $base64Imagen2 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($imagPath2))

                
            
                # Embeber las URLs de las im�genes en el cuerpo del correo
                $correoOpciones.Body =@"
                <!doctype html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<head>
  <title>
  </title>
  <!--[if !mso]><!-->
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <!--<![endif]-->
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style type="text/css">
    #outlook a {
      padding: 0;
    }

    body {
      margin: 0;
      padding: 0;
      -webkit-text-size-adjust: 100%;
      -ms-text-size-adjust: 100%;
    }

    table,
    td {
      border-collapse: collapse;
      mso-table-lspace: 0pt;
      mso-table-rspace: 0pt;
    }

    img {
      border: 0;
      height: auto;
      line-height: 100%;
      outline: none;
      text-decoration: none;
      -ms-interpolation-mode: bicubic;
    }

    p {
      display: block;
      margin: 13px 0;
    }
  </style>
  <!--[if mso]>
        <noscript>
        <xml>
        <o:OfficeDocumentSettings>
          <o:AllowPNG/>
          <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
        </xml>
        </noscript>
        <![endif]-->
  <!--[if lte mso 11]>
        <style type="text/css">
          .mj-outlook-group-fix { width:100% !important; }
        </style>
        <![endif]-->
  <!--[if !mso]><!-->
  <link href="https://fonts.googleapis.com/css?family=Droid+Sans:300,400,500,700" rel="stylesheet" type="text/css">
  <link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700" rel="stylesheet" type="text/css">
  <link href="https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700" rel="stylesheet" type="text/css">
  <style type="text/css">
    @import url(https://fonts.googleapis.com/css?family=Droid+Sans:300,400,500,700);
    @import url(https://fonts.googleapis.com/css?family=Roboto:300,400,500,700);
    @import url(https://fonts.googleapis.com/css?family=Ubuntu:300,400,500,700);
  </style>
  <!--<![endif]-->
  <style type="text/css">
    @media only screen and (min-width:600px) {
      .mj-column-per-100 {
        width: 100% !important;
        max-width: 100%;
      }
    }
  </style>
  <style media="screen and (min-width:600px)">
    .moz-text-html .mj-column-per-100 {
      width: 100% !important;
      max-width: 100%;
    }
  </style>
  <style type="text/css">
    @media only screen and (max-width:600px) {
      table.mj-full-width-mobile {
        width: 100% !important;
      }

      td.mj-full-width-mobile {
        width: auto !important;
      }
    }
  </style>
</head>

<body style="word-spacing:normal;background-color:#F7F7F7;">
  <div style="background-color:#F7F7F7;">
    <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
        <tbody>
          <tr>
            <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                  <tbody>
                    <tr>
                      <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                          <tbody>
                            <tr>
                              <td align="center" style="font-size:0px;padding:0px 0px 0px 0px;word-break:break-word;">
                                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;border-spacing:0px;">
                                  <tbody>
                                    <tr>
                                      <td style="width:800px;">
                                        <img alt="imagen" height="auto" src="data:image/png;base64,$base64Imagen1" style="border:0;display:block;outline:none;text-decoration:none;height:auto;width:100%;font-size:13px;" width="800" />
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><![endif]-->
    <!-- BODY -->
    <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
        <tbody>
          <tr>
            <td style="direction:ltr;font-size:0px;padding:20px 0;text-align:center;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;">
                        <div style="font-family:-apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen', 'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans','Helvetica Neue', sans-serif;font-size:14px;font-weight:400;line-height:18px;text-align:left;color:#000000;">
                          <p> Estimados colaboradores: </p>
                          <p>Informamos sobre una interrupcion temporal de sistemas en la <strong>sucursal $sucursal</strong></p>
                          <p>Nuestro equipo de TI ya esta trabajando para resolver este inconveniente con el proveedor y restablecer la normalidad en el menor tiempo posible.</p>
                          <p>En caso de emergencia, favor comunicarse al <strong>numero +56 $telefono.</strong></p>
                          <p>Los mantendremos informados.</p>
                          <p>Atentamente,</p>
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
        <tbody>
          <tr>
            <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                  <tbody>
                    <tr>
                      <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                          <tbody>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><![endif]-->
    <!-- END BODY -->
    <!--[if mso | IE]><table align="center" border="0" cellpadding="0" cellspacing="0" class="" style="width:800px;" width="800" bgcolor="#FFFFFF" ><tr><td style="line-height:0px;font-size:0px;mso-line-height-rule:exactly;"><![endif]-->
    <div style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;max-width:800px;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;">
        <tbody>
          <tr>
            <td style="border:none;direction:ltr;font-size:0px;padding:0px 0px 0px 0px;text-align:left;">
              <!--[if mso | IE]><table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td class="" style="vertical-align:top;width:800px;" ><![endif]-->
              <div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;">
                <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%">
                  <tbody>
                    <tr>
                      <td style="border:none;vertical-align:top;padding:0px 0px 0px 0px;">
                        <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="" width="100%">
                          <tbody>
                            <tr>
                              <td align="center" style="font-size:0px;padding:0px 0px 0px 0px;word-break:break-word;">
                                <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse;border-spacing:0px;">
                                  <tbody>
                                    <tr>
                                      <td style="width:800px;">
                                        <img alt="imagen" height="auto" src="data:image/png;base64,$base64Imagen2" style="border:0;display:block;outline:none;text-decoration:none;height:auto;width:100%;font-size:13px;" width="800" />
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <!--[if mso | IE]></td></tr></table><![endif]-->
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!--[if mso | IE]></td></tr></table><![endif]-->
  </div>
</body>

</html>
"@       
                # Enviar el correo electr�nico
                Send-MailMessage @correoOpciones -Bcc ($destinatarios+$correoSucursal) -ErrorAction Stop
                Write-Host "Correo electr�nico enviado a $($destinatarios+$correoSucursal -join ',')"
            
                # Estado anterior para comparar
                $row | Add-Member -MemberType NoteProperty -Name EstadoAnterior -Value "DOWN" -Force
            }
        }
    }

    # Duerme el script por una cantidad de segundos
    Start-Sleep -Seconds 1
}
