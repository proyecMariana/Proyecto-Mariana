//Metodo para crear archivo PDF de finalización de proceso
function crearPdf_Retiro(id = "43") {
  //This value should be the id of your document template that we created in the last step
  const googleDocTemplate = DriveApp.getFileById('1FMXEvRNZbM63BXALJcGPdj54t56OPVBNNsDozJOWUPc');
  // const googleDocTemplate = "1Uis10m1FKRSbtZw1xsnIhFMrNwreSdvy";

  //This value should be the id of the folder where you want your completed documents stored Documentos
  const destinationFolder = DriveApp.getFolderById('1fzQjz7Imh8CtjSYuYuKWasibxsfFhb1R')

  //This value should be the id of the folder where you want your completed documents stored PDF
  const destinationFolderPDF = DriveApp.getFolderById('1Z_Wv-ZyFJRCLwZdBaPyWAczzQlrqa9_X')
  //Here we store the sheet as a variable
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro Retiros')

  //--------------------------------------------------
  //Se guardar todos los registros de la lista para luego pasar a ser filtrados
  var idbus = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
  // se filtra por el id del cual se quiere convertir a pdf
  var posIndice = idbus.indexOf(id.toString().toLowerCase());
  //se toma le indice preciso de la busqueda para luego ubicarse y realizar los cambios 
  var rowNumberRet = posIndice === -1 ? 0 : posIndice + 2;
  //Obteniendo en un vector todos los campos de la busqueda, para iniciar el recorrido de cada uno de ellos con relacion al indice la busqueda
  var rows = sheet.getRange(rowNumberRet, 1, rowNumberRet, 38).getValues();
  // Logger.log('Logitud registros custIds3 :' + idbus + '\n Logitud encontrados posIndice:' + posIndice + '\n rowNumberRet:' + rowNumberRet + '\n rows-->:' + rows);
  //--------------------------------------------------
  rows.forEach(function (row, index) {
    if (index > 0) return;
    if (row[0].toString().trim() === id.toString().trim() && index === 0) {

     // Logger.log('\n Datos consulta ' + row[9] + '--' + row[10]);
      const copy = googleDocTemplate.makeCopy(`${row[9]}_${row[10]}_Retiro`, destinationFolder);
      //Once we have the copy, we then open it using the DocumentApp
      const doc = DocumentApp.openById(copy.getId());

      //----------------------------------------------

      //All of the content lives in the body, so we get that for editing
      const body = doc.getBody();
      //In this line we do some friendly date formatting, that may or may not work for you locale
      const friendlyDate = new Date(row[1]).toLocaleDateString();

      //In these lines, we replace our replacement tokens with values from our spreadsheet row

      //Datos para envio al documento a generar pdf
      body.replaceText('{{Apellidos Nombres}}', row[9]);
      body.replaceText('{{Codigo}}', row[10]);
      body.replaceText('{{Identificación}}', row[7]);
      body.replaceText('{{Programa académico}}', row[3]);
      body.replaceText('{{Semestre}}', row[5]);
      body.replaceText('{{Jornada}}', row[6]);
      body.replaceText('{{Periodo académico}}', row[2].toString());
      body.replaceText('{{Celular}}', row[12]);
      body.replaceText('{{Correo Electrónico}}', row[13]);
      body.replaceText('{{Obs.Prog}}', row[14]);
      body.replaceText('{{Obs.CC}}', row[19]);
      body.replaceText('{{Obs.AI}}', row[23]);
      body.replaceText('{{Obs.VA}}', row[27]);
      body.replaceText('{{Obs.VF}}', row[29]);
      body.replaceText('{{Obs.RC}}', row[33]);
      body.replaceText('{{EstadoRetiro}}', row[24]);
      body.replaceText('{{FechaInforme}}', new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }));
      body.replaceText('{{idRet}}', zfill(row[0], 6));
      body.replaceText('{{ConSinEfectos}}', row[16] === 'Con efectos' ? 'Con efectos académicos, los registros académicos se mantienen' : 'El registro académico no tiene efecto académicos y débe eliminarse');

      //###################################################################################################
      // Insertar imagen para firma de Retiro Estudiantil

      //Bloqueado mientras se baja tiempo de generacion de documento      

      var replaceTextToImage = function (body, searchText, image, width) {
        var next = body.findText(searchText);
        if (!next) return;
        var r = next.getElement();
        r.asText().setText("");
        var img = r.getParent().asParagraph().insertInlineImage(0, image);
        if (width && typeof width == "number") {
          var w = img.getWidth();
          var h = img.getHeight();
          img.setWidth(width);
          img.setHeight(width * h / w);
        }
        return next;
      };

     // Logger.log('Estado para imagen: ' + row[34]+' 37: '+ row[37]+' 38: '+row[38]);
      if (row[34] === "Cierre" && row[37]!='') {
        var replaceText1 = "RCX";
        var rutaver;
        if (row[37] != '') {
          rutaver = row[37].split('/')[5];
        } else {
          rutaver = '';
        }
        var imageFileId1 = rutaver;
        var body1 = body;
        var image1 = DriveApp.getFileById(imageFileId1).getBlob();
        do {
          var next1 = replaceTextToImage(body1, replaceText1, image1, 600);
        } while (next1);
      }// Fin if
  
      //###################################################################################################   

      //Hacemos nuestros cambios permanentes guardando y cerrando el documento
      doc.saveAndClose();
      //Logger.log('Documento creado y cerrando conexion');
      //Convirtiendo a PDF para su guardado
      const pdfContentBlob = doc.getAs(MimeType.PDF);
      const copy1 = destinationFolderPDF.createFile(pdfContentBlob).setName(`${row[9]}_${row[10]}_Retiro`);
      //Store the url of our new document pdf in a variable
      const url = copy1.getUrl();
      Logger.log('Obteniendo URL para fijar archivo: ' + url);
      //Write that value back to the 'Document Link' column in the spreadsheet. 
      sheet.getRange(rowNumberRet, 36, 1, 1).setValues([[
        url
      ]]);
      //fin validacion para generar documento
    }//fin if
  })
  // Fin foreach recorrido 
}//Fin crear pdf de retiro

//-----------------------------------------------------------------------------------------------
//Metodo para crear archivo PDF para envio al estudiante
function crearPdf_RetiroEstudiante(id = "43") {
  //This value should be the id of your document template that we created in the last step
  //Realizar plantilla para cambiar el pdf de envio al estudiante
  const googleDocTemplate = DriveApp.getFileById('1e6MsyncBfHKpwIIyt3TKbahfdzSoJd5d9obvkoFMiwE');
  // const googleDocTemplate = "1Uis10m1FKRSbtZw1xsnIhFMrNwreSdvy";

  //This value should be the id of the folder where you want your completed documents stored Documentos
  const destinationFolder = DriveApp.getFolderById('1fzQjz7Imh8CtjSYuYuKWasibxsfFhb1R')

  //This value should be the id of the folder where you want your completed documents stored PDF
  const destinationFolderPDF = DriveApp.getFolderById('1Z_Wv-ZyFJRCLwZdBaPyWAczzQlrqa9_X')
  //Here we store the sheet as a variable
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro Retiros')

  //--------------------------------------------------
  //Se guardar todos los registros de la lista para luego pasar a ser filtrados
  var idbus = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
  // se filtra por el id del cual se quiere convertir a pdf
  var posIndice = idbus.indexOf(id.toString().toLowerCase());
  //se toma le indice preciso de la busqueda para luego ubicarse y realizar los cambios 
  var rowNumberRet = posIndice === -1 ? 0 : posIndice + 2;
  //Obteniendo en un vector todos los campos de la busqueda, para iniciar el recorrido de cada uno de ellos con relacion al indice la busqueda
  var rows = sheet.getRange(rowNumberRet, 1, rowNumberRet, 36).getValues();
  // Logger.log('Logitud registros custIds3 :' + idbus + '\n Logitud encontrados posIndice:' + posIndice + '\n rowNumberRet:' + rowNumberRet + '\n rows-->:' + rows);
  //--------------------------------------------------
  rows.forEach(function (row, index) {
    if (index > 0) return;
    if (row[0].toString().trim() === id.toString().trim() && index === 0) {

     // Logger.log('\n Datos consulta ' + row[9] + '--' + row[10]);
      const copy = googleDocTemplate.makeCopy(`${row[9]}_${row[10]}_Retiro_Estudiante`, destinationFolder);
      //Once we have the copy, we then open it using the DocumentApp
      const doc = DocumentApp.openById(copy.getId());

      //----------------------------------------------

      //All of the content lives in the body, so we get that for editing
      const body = doc.getBody();
      //In this line we do some friendly date formatting, that may or may not work for you locale
      const friendlyDate = new Date(row[1]).toLocaleDateString();

      //In these lines, we replace our replacement tokens with values from our spreadsheet row

      //Datos para envio al documento a generar pdf
      body.replaceText('{{Apellidos Nombres}}', row[9]);
      body.replaceText('{{Codigo}}', row[10]);
      body.replaceText('{{Identificación}}', row[7]);
      body.replaceText('{{Programa académico}}', row[3]);
      body.replaceText('{{Semestre}}', row[5]);
      body.replaceText('{{Jornada}}', row[6]);
      body.replaceText('{{Periodo académico}}', row[2].toString());
      body.replaceText('{{Celular}}', row[12]);
      body.replaceText('{{Correo Electrónico}}', row[13]);
      body.replaceText('{{Obs.Prog}}', row[14]);
      body.replaceText('{{Obs.CC}}', row[19]);
      body.replaceText('{{Obs.AI}}', row[23]);
      body.replaceText('{{Obs.VA}}', row[27]);
      body.replaceText('{{Obs.VF}}', row[29]);
      body.replaceText('{{Obs.RC}}', row[33]);
      body.replaceText('{{EstadoRetiro}}', row[24]);
      body.replaceText('{{FechaInforme}}', new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }));
      body.replaceText('{{idRet}}', zfill(row[0], 6));
      body.replaceText('{{ConSinEfectos}}', row[16] === 'Con efectos' ? 'Con efectos académicos, los registros académicos se mantienen' : 'El registro académico no tiene efecto académicos y débe eliminarse');

      //###################################################################################################
      // Insertar imagen para firma de Retiro Estudiantil
      //Bloqueado mientras se baja tiempo de generacion de documento      
      /*
            var replaceTextToImage = function (body, searchText, image, width) {
              var next = body.findText(searchText);
              if (!next) return;
              var r = next.getElement();
              r.asText().setText("");
              var img = r.getParent().asParagraph().insertInlineImage(0, image);
              if (width && typeof width == "number") {
                var w = img.getWidth();
                var h = img.getHeight();
                img.setWidth(width);
                img.setHeight(width * h / w);
              }
              return next;
            };
            Logger.log('Estado PA: ' + row[18]);
            if (row[17] === "Completado") {
              var replaceText1 = "PA";
              var imageFileId1 = "1jXrWSf4iZzB3_bkk4kgnpvOINZ-Lt_dv";
              var body1 = body;
              var image1 = DriveApp.getFileById(imageFileId1).getBlob();
              do {
                var next1 = replaceTextToImage(body1, replaceText1, image1, 150);
              } while (next1);
            }// Fin if
            if (row[22] === "Completado") {
              var replaceText2 = "CC";
              var imageFileId2 = "1Rt3qsVnPyVXBXtQjEZaRIYTthVPQ98Fv";
              var body2 = body;
              var image2 = DriveApp.getFileById(imageFileId2).getBlob();
              do {
                var next2 = replaceTextToImage(body2, replaceText2, image2, 150);
              } while (next2);
            }
            if (row[25] === "Completado") {
              var replaceText3 = "AI";
              var imageFileId3 = "1jXrWSf4iZzB3_bkk4kgnpvOINZ-Lt_dv";
              var body3 = body;
              var image3 = DriveApp.getFileById(imageFileId3).getBlob();
              do {
                var next3 = replaceTextToImage(body3, replaceText3, image3, 150);
              } while (next3);
            }
            if (row[28] === "Completado") {
              var replaceText4 = "VA";
              var imageFileId4 = "1GS3m_BGAF8qhKGtaGRaWi1xXj7t5i2wx";
              var body4 = body;
              var image4 = DriveApp.getFileById(imageFileId4).getBlob();
              do {
                var next4 = replaceTextToImage(body4, replaceText4, image4, 150);
              } while (next4);
            }
            if (row[32] === "Completado") {
              var replaceText5 = "VF";
              var imageFileId5 = "1jj8StGNe-PggqxRAtsjEltZFmW72HsW9";
              var body5 = body;
              var image5 = DriveApp.getFileById(imageFileId5).getBlob();
              do {
                var next5 = replaceTextToImage(body5, replaceText5, image5, 150);
              } while (next5);
            }
      
            if (row[34] === "Cierre") {
              var replaceText6 = "RC";
              var imageFileId6 = "1YtobTOSOAUoQz7HhaXTw5SQkpK_zBd5e";
              var body6 = body;
              var image6 = DriveApp.getFileById(imageFileId6).getBlob();
              do {
                var next6 = replaceTextToImage(body6, replaceText6, image6, 150);
              } while (next6);
            }
      */ //Bloqueado mientras se baja tiempo de generacion de documento

      // Fin de insertar imagen para Paz y Salvo  
      //###################################################################################################   

      //Hacemos nuestros cambios permanentes guardando y cerrando el documento
      doc.saveAndClose();
      //Logger.log('Documento creado y cerrando conexion');
      //Convirtiendo a PDF para su guardado
      const pdfContentBlob = doc.getAs(MimeType.PDF);
      const copy1 = destinationFolderPDF.createFile(pdfContentBlob).setName(`${row[9]}_${row[10]}_Retiro_Estudiante`);
      //Store the url of our new document pdf in a variable
      const url = copy1.getUrl();
     // Logger.log('Obteniendo URL para fijar archivo: ' + url);
      //Write that value back to the 'Document Link' column in the spreadsheet. 
      sheet.getRange(rowNumberRet, 37, 1, 1).setValues([[
        url
      ]]);
      //fin validacion para generar documento
    }//fin if
  })
  // Fin foreach recorrido 
}//Fin crear pdf de retiro

//Metodo para crear archivo PDF de finalización de proceso
function crearPdf_Retiro(id = "43") {
  //This value should be the id of your document template that we created in the last step
  const googleDocTemplate = DriveApp.getFileById('1FMXEvRNZbM63BXALJcGPdj54t56OPVBNNsDozJOWUPc');
  // const googleDocTemplate = "1Uis10m1FKRSbtZw1xsnIhFMrNwreSdvy";

  //This value should be the id of the folder where you want your completed documents stored Documentos
  const destinationFolder = DriveApp.getFolderById('1fzQjz7Imh8CtjSYuYuKWasibxsfFhb1R')

  //This value should be the id of the folder where you want your completed documents stored PDF
  const destinationFolderPDF = DriveApp.getFolderById('1Z_Wv-ZyFJRCLwZdBaPyWAczzQlrqa9_X')
  //Here we store the sheet as a variable
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro Retiros')

  //--------------------------------------------------
  //Se guardar todos los registros de la lista para luego pasar a ser filtrados
  var idbus = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
  // se filtra por el id del cual se quiere convertir a pdf
  var posIndice = idbus.indexOf(id.toString().toLowerCase());
  //se toma le indice preciso de la busqueda para luego ubicarse y realizar los cambios 
  var rowNumberRet = posIndice === -1 ? 0 : posIndice + 2;
  //Obteniendo en un vector todos los campos de la busqueda, para iniciar el recorrido de cada uno de ellos con relacion al indice la busqueda
  var rows = sheet.getRange(rowNumberRet, 1, rowNumberRet, 38).getValues();
  // Logger.log('Logitud registros custIds3 :' + idbus + '\n Logitud encontrados posIndice:' + posIndice + '\n rowNumberRet:' + rowNumberRet + '\n rows-->:' + rows);
  //--------------------------------------------------
  rows.forEach(function (row, index) {
    if (index > 0) return;
    if (row[0].toString().trim() === id.toString().trim() && index === 0) {

     // Logger.log('\n Datos consulta ' + row[9] + '--' + row[10]);
      const copy = googleDocTemplate.makeCopy(`${row[9]}_${row[10]}_Retiro`, destinationFolder);
      //Once we have the copy, we then open it using the DocumentApp
      const doc = DocumentApp.openById(copy.getId());

      //----------------------------------------------

      //All of the content lives in the body, so we get that for editing
      const body = doc.getBody();
      //In this line we do some friendly date formatting, that may or may not work for you locale
      const friendlyDate = new Date(row[1]).toLocaleDateString();

      //In these lines, we replace our replacement tokens with values from our spreadsheet row

      //Datos para envio al documento a generar pdf
      body.replaceText('{{Apellidos Nombres}}', row[9]);
      body.replaceText('{{Codigo}}', row[10]);
      body.replaceText('{{Identificación}}', row[7]);
      body.replaceText('{{Programa académico}}', row[3]);
      body.replaceText('{{Semestre}}', row[5]);
      body.replaceText('{{Jornada}}', row[6]);
      body.replaceText('{{Periodo académico}}', row[2].toString());
      body.replaceText('{{Celular}}', row[12]);
      body.replaceText('{{Correo Electrónico}}', row[13]);
      body.replaceText('{{Obs.Prog}}', row[14]);
      body.replaceText('{{Obs.CC}}', row[19]);
      body.replaceText('{{Obs.AI}}', row[23]);
      body.replaceText('{{Obs.VA}}', row[27]);
      body.replaceText('{{Obs.VF}}', row[29]);
      body.replaceText('{{Obs.RC}}', row[33]);
      body.replaceText('{{EstadoRetiro}}', row[24]);
      body.replaceText('{{FechaInforme}}', new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }));
      body.replaceText('{{idRet}}', zfill(row[0], 6));
      body.replaceText('{{ConSinEfectos}}', row[16] === 'Con efectos' ? 'Con efectos académicos, los registros académicos se mantienen' : 'El registro académico no tiene efecto académicos y débe eliminarse');

      //###################################################################################################
      // Insertar imagen para firma de Retiro Estudiantil

      //Bloqueado mientras se baja tiempo de generacion de documento      

      var replaceTextToImage = function (body, searchText, image, width) {
        var next = body.findText(searchText);
        if (!next) return;
        var r = next.getElement();
        r.asText().setText("");
        var img = r.getParent().asParagraph().insertInlineImage(0, image);
        if (width && typeof width == "number") {
          var w = img.getWidth();
          var h = img.getHeight();
          img.setWidth(width);
          img.setHeight(width * h / w);
        }
        return next;
      };

     // Logger.log('Estado para imagen: ' + row[34]+' 37: '+ row[37]+' 38: '+row[38]);
      if (row[34] === "Cierre" && row[37]!='') {
        var replaceText1 = "RCX";
        var rutaver;
        if (row[37] != '') {
          rutaver = row[37].split('/')[5];
        } else {
          rutaver = '';
        }
        var imageFileId1 = rutaver;
        var body1 = body;
        var image1 = DriveApp.getFileById(imageFileId1).getBlob();
        do {
          var next1 = replaceTextToImage(body1, replaceText1, image1, 600);
        } while (next1);
      }// Fin if

      //###################################################################################################   

      //Hacemos nuestros cambios permanentes guardando y cerrando el documento
      doc.saveAndClose();
      //Logger.log('Documento creado y cerrando conexion');
      //Convirtiendo a PDF para su guardado
      const pdfContentBlob = doc.getAs(MimeType.PDF);
      const copy1 = destinationFolderPDF.createFile(pdfContentBlob).setName(`${row[9]}_${row[10]}_Retiro`);
      //Store the url of our new document pdf in a variable
      const url = copy1.getUrl();
      Logger.log('Obteniendo URL para fijar archivo: ' + url);
      //Write that value back to the 'Document Link' column in the spreadsheet. 
      sheet.getRange(rowNumberRet, 36, 1, 1).setValues([[
        url
      ]]);
      //fin validacion para generar documento
    }//fin if
  })
  // Fin foreach recorrido 
}//Fin crear pdf de retiro

//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------
//Metodo para crear archivo PDF de solicitud de pago anticipado
function crearPdf_Pago_Anticipado(id = "43", vrord="0") {
  //This value should be the id of your document template that we created in the last step
  //Realizar plantilla para cambiar el pdf de envio al estudiante
  const googleDocTemplate = DriveApp.getFileById('1tMi0BPgYSS6EBdto2mz6EQAckV8wa9ju37YSXy6rFuc');

  //This value should be the id of the folder where you want your completed documents stored Documentos
  const destinationFolder = DriveApp.getFolderById('1b2WTEprwyrvFuLJBKIGFwCi7vNCc8SYw')

  //This value should be the id of the folder where you want your completed documents stored PDF
  const destinationFolderPDF = DriveApp.getFolderById('1ujgtHT6tYFppBsivjTcVY1kQWdSZtA4Y')
  //Here we store the sheet as a variable
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GenRec')

  //--------------------------------------------------
  //Se guardar todos los registros de la lista para luego pasar a ser filtrados
  var idbus = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
  // se filtra por el id del cual se quiere convertir a pdf
  var posIndice = idbus.indexOf(id.toString().toLowerCase());
  //se toma le indice preciso de la busqueda para luego ubicarse y realizar los cambios 
  var rowNumberRet = posIndice === -1 ? 0 : posIndice + 2;
  //Obteniendo en un vector todos los campos de la busqueda, para iniciar el recorrido de cada uno de ellos con relacion al indice la busqueda
  var rows = sheet.getRange(rowNumberRet, 1, rowNumberRet, 23).getValues();
  // Logger.log('Logitud registros custIds3 :' + idbus + '\n Logitud encontrados posIndice:' + posIndice + '\n rowNumberRet:' + rowNumberRet + '\n rows-->:' + rows);
  //--------------------------------------------------
  rows.forEach(function (row, index) {
    if (index > 0) return;
    if (row[0].toString().trim() === id.toString().trim() && index === 0) {

      //Logger.log('\n Datos consulta ' + row[7] + '--' + row[0]);
      const copy = googleDocTemplate.makeCopy(`${row[7]}_${row[0]}_pago_anticipado`, destinationFolder);
      //Once we have the copy, we then open it using the DocumentApp
      const doc = DocumentApp.openById(copy.getId());

      //----------------------------------------------

      //All of the content lives in the body, so we get that for editing
      const body = doc.getBody();
      //In this line we do some friendly date formatting, that may or may not work for you locale
      const friendlyDate = new Date(row[1]).toLocaleDateString();

      //In these lines, we replace our replacement tokens with values from our spreadsheet row

      //Datos para envio al documento a generar pdf
      body.replaceText('{{fecha_gen}}', row[1].toLocaleString());
      body.replaceText('{{Apellidos Nombres}}', row[7]);
      body.replaceText('{{Codigo}}', row[8]);
      body.replaceText('{{identificacion}}', row[6]);
      body.replaceText('{{Programa académico}}', row[3]);
      body.replaceText('{{Semestre}}', row[5]); //revisar
      body.replaceText('{{Jornada}}', row[5]);
      body.replaceText('{{Periodo académico}}', row[2].toString());
      body.replaceText('{{pago_5}}', formatterPeso.format(row[11]));
      body.replaceText('{{pago_3}}', formatterPeso.format(row[12]));
      body.replaceText('{{pago_2}}', formatterPeso.format(row[13]));
      body.replaceText('{{pago_ord}}', formatterPeso.format(vrord));
      body.replaceText('{{idRec}}', zfill(row[0], 6));
      body.replaceText('{{descomp}}', row[21]+' Semestre (Pago Anticipado)');  
      body.replaceText('{{creditos}}', row[22]);  
      
      //Queda pendiente para realizar pruebas mas a fondo
      /*
      var resp5 = UrlFetchApp.fetch(row[14]);
      body.replaceText('{{barras_5}}', body.insertImage(0,resp5.getBlob())); 
      var resp3 = UrlFetchApp.fetch(row[16]);
      body.replaceText('{{barras_3}}', body.insertImage(0,resp3.getBlob())); 
      var resp2 = UrlFetchApp.fetch(row[18]);
      body.replaceText('{{barras_2}}', body.insertImage(0,resp2.getBlob())); */
      //###################################################################################################

      //Hacemos nuestros cambios permanentes guardando y cerrando el documento
      doc.saveAndClose();
      //Logger.log('Documento creado y cerrando conexion');
      //Convirtiendo a PDF para su guardado
      const pdfContentBlob = doc.getAs(MimeType.PDF);
      const copy1 = destinationFolderPDF.createFile(pdfContentBlob).setName(`${row[7]}_${row[0]}_Pago_Anticipado`);
      //Store the url of our new document pdf in a variable
      const url = copy1.getUrl();
      //Logger.log('Obteniendo URL para fijar archivo: ' + url);
      //Write that value back to the 'Document Link' column in the spreadsheet. 
      if(url !='')  sheet.getRange(rowNumberRet, 24, 1, 1).setValues([[url]]);
      else swal('Error en busqueda','No se genero documento pdf, verifique perdida de sesión de correo electrónico','error');
      //fin validacion para generar documento
    }//fin if
  })
  // Fin foreach recorrido 
}//Fin crear pdf de retiro

//Formatear valor numerico para mostrar en tipo moneda
    const formatterPeso = new Intl.NumberFormat('es-CO', {
       style: 'currency',
       currency: 'COP',
       minimumFractionDigits: 0
     })
