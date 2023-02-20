function envioCorreo(nivRetiro, idRet, progret, apenom, codigo, origen, destino, emaildestino) {

  // alert('Programa:' + programa + '\n Correo: ' + emailAddress + '\n Asunto: ' + asunto + '\n Observacion: ' + nivRetiro);
  //'CRÉDITO Y CARTERA', 'gwsanchez@UMARIANA.edu.co  ', 'Solicitud de Verificación de Retiro',

  var emailOrigen = Session.getActiveUser().getEmail();

  var idRetmail = idRet
  var origen_m = origen;
  var destino_m = destino;
  var emailEnvio = emaildestino;
  // Logger.log('Datos Usuario unico: -->' + origen_m + '\n Destino:' + destino_m + '\n Correo Destino:' + emailEnvio+' ->' +progret+' \n nivRetiro'+nivRetiro);
  var testmessagebody;
  //get the html content from file.
  //console.log('Antes de irse por las opciones de decision de envio de email:'+nivRetiro.toString());
  switch (nivRetiro.toString()) {
    case '7':
      console.log('Entro en anulacion');
      testmessagebody = HtmlService.createHtmlOutputFromFile('mail_envio_3').getContent()
      break;
    case '6':
    console.log('Entro en Finalizacion retiro');
      testmessagebody = HtmlService.createHtmlOutputFromFile('mail_envio_2').getContent()
      break;
    case 'Rec':
      console.log('Envio Recibo');
      testmessagebody = HtmlService.createHtmlOutputFromFile('mail_envio_Rec_Ant').getContent()
      break;      
    default:
      console.log('Entro en opcion ciclica');
      testmessagebody = HtmlService.createHtmlOutputFromFile('mail_envio_1').getContent()
      break;
  }

  switch (nivRetiro.toString()) {
    case '1': //Envio Crédito y Cartera / Programa
      var rutaAcceso = rutaGlobal + '?p=4&d=' + idRet
      var asunto = 'Solicitud de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });

      break;
    case '2': //Permanencia Estudiantial / Credito y cartera
      var rutaAcceso = rutaGlobal + '?p=5&d=' + idRet
      var asunto = 'Solicitud de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    case '3': //Vicerrectoria Académica / Permanencia estudiantil
      var rutaAcceso = rutaGlobal + '?p=6&d=' + idRet
      var asunto = 'Solicitud de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    case '4': //Vicerrectoria Administrativa y Financiera / Vicerrectoría Académica
      var rutaAcceso = rutaGlobal + '?p=7&d=' + idRet
      var asunto = 'Solicitud de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    case '5': //Registro y Control Académico / Vicerrectoría Administrativa y Financiera
      var rutaAcceso = rutaGlobal + '?p=8&d=' + idRet
      var asunto = 'Solicitud de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    case '6': // Registro y Control Académico / Programa
    //console.log('Revisando la cadena de aquien se envia:'+emailEnvio+'Destino envio:'+destino_m);
      var separarDestino = destino_m.split('-');
      testmessagebody = HtmlService.createHtmlOutputFromFile('mail_envio_2').getContent()
      var rutaAcceso = rutaGlobal + '?p=11&d=' + idRet
      var asunto = 'Finalización de Verificación de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", separarDestino[0])
        .replace("{destinoC}", separarDestino[1])        
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    case '7': // Anulando proceso de retiro
      var rutaAcceso = rutaGlobal + '?p=8&d=' + idRet
      var asunto = 'Anulando solicitud de Retiro No.' + idRet;
      var email_body = testmessagebody.replace("{areaEnv}", destino_m)
        .replace("{asunto}", asunto)
        .replace("{progestudiante}", progret)
        .replace("{rutaAcceso}", rutaAcceso)
        .replace("{destino}", destino_m)
        .replace("{origen}", origen_m)
        .replace("{apenom}", apenom)
        .replace("{codigo}", codigo)

      MailApp.sendEmail({
        cc: emailOrigen, //cc manager
        to: emailEnvio,
        subject: asunto,
        htmlBody: email_body
      });
      break;
    default:
      break;
  }
}
