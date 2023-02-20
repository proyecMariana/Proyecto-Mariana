// Global variables
var HE = SpreadsheetApp.openById('1xqgAU0SvtZqdDxPEpJzvk29q36is1TLlUT_1JN_HHxY');
var sheetBD = HE.getSheetByName('Listas');
const rutaGlobal = "https://script.google.com/a/macros/umariana.edu.co/s/AKfycbz-GKU3jigLd6ImJ2Xt03Zriq9Q50Cj8p7kUhpUXdHp/dev";
//IMPLEMENTACION CINCO
//const rutaGlobal = "https://script.google.com/a/macros/UMARIANA.edu.co/s/AKfycbwVdQz89dP4avRwrTzIj8zcchyCF2w_YBhDS97yQvkj660f9bJ2Jkk2A87cRK4Gd6Cn-Q/exec";

function abrirLink(){
  var html = HtmlService.createHtmlOutput('<html>'
  +'<script>'+
         "var urlToOpen = 'https://script.google.com/a/macros/umariana.edu.co/s/AKfycbz-GKU3jigLd6ImJ2Xt03Zriq9Q50Cj8p7kUhpUXdHp/dev';"+
         "var winRef = window.open(urlToOpen);"+
         "google.script.host.close();"
  +'</script>'
  +'</html>')
  .setWidth(90).setHeight(1); 
  SpreadsheetApp.getUi().showModalDialog(html, "Abriendo ...");
}


function doGet(e) {
  //-----------------------------------------------------
  const permitirAcceso = searchUser();
  if (permitirAcceso === true) {

    let params = e.parameter;
    let pantalla = params.p;
    //console.log('Que es esto doGet : ' + params.p + '\n params d: ' + params.d);

    switch (params.p) {
      case "1":  //Formulario de ingreso de información para las secretarias
        //Llamado a lista dinamica programa
        //var data = sheetBD.getDataRange().getValues();
        //Esta linea permite tener la informacion de excel en la variable data pero iniciando desde la fila 2 y tomando las 7 columnas del archivo
        var data = sheetBD.getRange(2, 1, sheetBD.getLastRow() - 1, 7).getValues();
        var programas1 = [];
        data.forEach(row => {
          if (programas1.indexOf(row[0]) == -1) {
            programas1.push(row[0]);
          }
        });
        //Llamado a lista dinamica periodo
        var peracad = [];
        data.forEach(row => {
          if (peracad.indexOf(row[2]) == -1) {
            peracad.push(row[2]);
          }
        });

        //Llamado a lista dinamica tipo documento
        var tipodoc = [];
        data.forEach(row => {
          if (tipodoc.indexOf(row[6]) == -1) {
            tipodoc.push(row[6]);
          }
        });
        var template1 = HtmlService.createTemplateFromFile('RegistroPA'); //creando una plantilla desde un archivo
        template1.programas1 = programas1;
        template1.peracad = peracad;
        //template1.semdinamico = semdinamico;
        template1.tipodoc = tipodoc;
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "2": //formulario de filtro de casos por cada dependencia verificadora
        var template1 = HtmlService.createTemplateFromFile("RegistroTerminado")
        //template1.pubUrl = rutaGlobal+"";
        var output = template1.evaluate();
        break;
      case "3": //formulario de Busqueda
        var template1 = HtmlService.createTemplateFromFile("BuscarRetiros");
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "4": //formulario de complementacion Credito y Cartera
        //console.log('Dato recibido: ' + params.d);
        var id = params.d;
        const ss5 = SpreadsheetApp.getActiveSpreadsheet();
        const ws5 = ss5.getSheetByName("Registro Retiros");
        const custIds5 = ws5.getRange(2, 1, ws5.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
        const posIndex5 = custIds5.indexOf(id.toString().toLowerCase());
        const rowNumber5 = posIndex5 === -1 ? 0 : posIndex5 + 2;
        const customerInfo5 = ws5.getRange(rowNumber5, 1, 1, 39).getValues()[0];
        //Verificando si se ha cargado imagen de notas
        var rutaver;
        if (customerInfo5[37] != '') {
          rutaver = "https://drive.google.com/uc?export=view&id=" + customerInfo5[37].split('/')[5];
        } else {
          rutaver = '';
        }

        //-------------------------------------------------------------------------
        var template1 = HtmlService.createTemplateFromFile('RegistroCreditoCartera'); //creando una plantilla desde un archivo
        template1.customerInfo = customerInfo5;
        template1.imgnotas = rutaver;
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "5": //formulario de complementacion Acompañamiento Integral
        //Llamado a lista dinamica programa
        //var data = sheetBD.getDataRange().getValues();
        //Esta linea permite tener la informacion de excel en la variable data pero iniciando desde la fila 2 y tomando las 7 columnas del archivo
        var data = sheetBD.getRange(8, 1, sheetBD.getLastRow() - 1, 8).getValues();
        var estadoRetiro = [];
        data.forEach(row => {
          if (estadoRetiro.indexOf(row[7]) == -1) {
            estadoRetiro.push(row[7]);
          }
        });
        //console.log('Dato recibido: ' + params.d);
        var id = params.d;
        const ss6 = SpreadsheetApp.getActiveSpreadsheet();
        const ws6 = ss6.getSheetByName("Registro Retiros");
        const custIds6 = ws6.getRange(2, 1, ws6.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
        const posIndex6 = custIds6.indexOf(id.toString().toLowerCase());
        const rowNumber6 = posIndex6 === -1 ? 0 : posIndex6 + 2;
        const customerInfo6 = ws6.getRange(rowNumber6, 1, 1, 39).getValues()[0];
        //Verificando si se ha cargado imagen de notas
        var rutaver;
        if (customerInfo6[37] != '') {
          rutaver = "https://drive.google.com/uc?export=view&id=" + customerInfo6[37].split('/')[5];
        } else {
          rutaver = '';
        }
        //-------------------------------------------------------------------------
        var template1 = HtmlService.createTemplateFromFile('RegistroAI'); //creando una plantilla desde un archivo
        template1.customerInfo = customerInfo6;
        template1.imgnotas = rutaver;
        template1.estadoRetiro = estadoRetiro;
        template1.pubUrl = rutaGlobal + "?p=5";
        template1.urlG = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "6": //formulario de complementacion Vice. Académica
        //console.log('Dato recibido: ' + params.d);
        var id = params.d;
        const ss7 = SpreadsheetApp.getActiveSpreadsheet();
        const ws7 = ss7.getSheetByName("Registro Retiros");
        const custIds7 = ws7.getRange(2, 1, ws7.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
        const posIndex7 = custIds7.indexOf(id.toString().toLowerCase());
        const rowNumber7 = posIndex7 === -1 ? 0 : posIndex7 + 2;
        //-- Dejar la posicion mas 1
        const customerInfo7 = ws7.getRange(rowNumber7, 1, 1, 39).getValues()[0];
        //Verificando si se ha cargado imagen de notas
        var rutaver;
        if (customerInfo7[37] != '') {
          rutaver = "https://drive.google.com/uc?export=view&id=" + customerInfo7[37].split('/')[5];
        } else {
          rutaver = '';
        }
        //-------------------------------------------------------------------------
        var template1 = HtmlService.createTemplateFromFile('RegistroVA'); //creando una plantilla desde un archivo
        template1.customerInfo = customerInfo7;
        template1.imgnotas = rutaver;
        template1.pubUrl = rutaGlobal + "?p=6";
        template1.urlG = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "7": //formulario de complementacion Vice Admin.
        //console.log('Dato recibido: ' + params.d);
        //Llamado a lista Concepto Final
        var data = sheetBD.getDataRange().getValues();
        var concepSaldo = [];
        data.forEach(row => {
          if (concepSaldo.indexOf(row[5]) == -1) {
            concepSaldo.push(row[5]);
          }
        });
        var id = params.d;
        const ss8 = SpreadsheetApp.getActiveSpreadsheet();
        const ws8 = ss8.getSheetByName("Registro Retiros");
        const custIds8 = ws8.getRange(2, 1, ws8.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
        const posIndex8 = custIds8.indexOf(id.toString().toLowerCase());
        const rowNumber8 = posIndex8 === -1 ? 0 : posIndex8 + 2;
        const customerInfo8 = ws8.getRange(rowNumber8, 1, 1, 39).getValues()[0];
        //Verificando si se ha cargado imagen de notas
        var rutaver;
        if (customerInfo8[37] != '') {
          rutaver = "https://drive.google.com/uc?export=view&id=" + customerInfo8[37].split('/')[5];
        } else {
          rutaver = '';
        }
        //-------------------------------------------------------------------------
        var template1 = HtmlService.createTemplateFromFile('RegistroVF'); //creando una plantilla desde un archivo
        template1.customerInfo = customerInfo8;
        template1.imgnotas = rutaver;
        template1.concepSaldo = concepSaldo;
        template1.pubUrl = rutaGlobal + "?p=7";
        template1.urlG = rutaGlobal;
        var output = template1.evaluate();
        break;
      case "8": //formulario de complementacion Registro y Control.
        //console.log('Dato recibido: ' + params.d);
        var id = params.d;
        const ss9 = SpreadsheetApp.getActiveSpreadsheet();
        const ws9 = ss9.getSheetByName("Registro Retiros");
        const custIds9 = ws9.getRange(2, 1, ws9.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
        const posIndex9 = custIds9.indexOf(id.toString().toLowerCase());
        const rowNumber9 = posIndex9 === -1 ? 0 : posIndex9 + 2;
        const customerInfo9 = ws9.getRange(rowNumber9, 1, 1, 39).getValues()[0];
        //Verificando si se ha cargado imagen de notas
        var rutaver;
        if (customerInfo9[37] != '') {
          rutaver = "https://drive.google.com/uc?export=view&id=" + customerInfo9[37].split('/')[5];
        } else {
          rutaver = '';
        }
        //-------------------------------------------------------------------------
        var template1 = HtmlService.createTemplateFromFile('RegistroRC'); //creando una plantilla desde un archivo
        template1.customerInfo = customerInfo9;
        template1.pubUrl = rutaGlobal + "?p=8";
        template1.urlG = rutaGlobal;
        template1.imgnotas = rutaver;
        var output = template1.evaluate();
        break;
      case "9": //formulario de Busqueda de Retiros
        var output = HtmlService.createTemplateFromFile("BuscarRetiros")
          .getRawContent
          .evaluate();
        break;
      case "10": //Alerta sin Documento PDF
        var output = HtmlService.createTemplateFromFile("SinDocumentoPDF")
          .evaluate();
        break;
      case "11": //formulario de Busqueda de Retiros para las dependencias diferentes a los programas
        var id = params.d;
        //console.log('Antes de validar:' + id);
        //var id = 43;
        //console.log('Valor de entrada hoy:' + typeof idcod);
        if (id === undefined || id === null) {
          console.log('Entro por variable no definida:' + id);
          var template1 = HtmlService.createTemplateFromFile("EstadosRetiros");
          template1.pubUrl = rutaGlobal;
          template1.rutaCierre = rutaGlobal;
          template1.codigo = undefined;
          var output = template1.evaluate();
        }
        else {
          console.log('Valor de entrada porque id no esta vacio:' + id);
          //Buscando codigo del estudiante, para enviarlo al input de buscar registro y se active la busqueda automatica
          const ss9 = SpreadsheetApp.getActiveSpreadsheet();
          const ws9 = ss9.getSheetByName("Registro Retiros");
          const custIds9 = ws9.getRange(2, 1, ws9.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
          const posIndex9 = custIds9.indexOf(id.toString().toLowerCase());
          const rowNumber9 = posIndex9 === -1 ? 0 : posIndex9 + 2;
          const customerInfo9 = ws9.getRange(rowNumber9, 1, 1, 11).getValues()[0];
          //Fin del proceso de buscar codigo
          //console.log('customerInfox:' + customerInfox[10]);
          var tmp = HtmlService.createTemplateFromFile("EstadosRetiros");
          tmp.pubUrl = rutaGlobal + "?p=11&d=" + id;
          tmp.codigo = customerInfo9[10];
          tmp.rutaCierre = rutaGlobal;
          var output = tmp.evaluate();
        }
        break;
      case "12": //Formulario guia telefónica
        var template = HtmlService.createTemplateFromFile("Bus_telefono");
        template.pubUrl = rutaGlobal;
        var output = template.evaluate();
        break;
      case "13": //Formulario tarifas matricula
        var template = HtmlService.createTemplateFromFile("Bus_tarifas_matricula");
        template.pubUrl = rutaGlobal;
        var output = template.evaluate();
        break; 
      case "14": //Formulario calculo pago descuento parcial
        var data = sheetBD.getRange(2, 1, sheetBD.getLastRow() - 1, 7).getValues();
        var programas1 = [];
        data.forEach(row => {
          if (programas1.indexOf(row[0]) == -1) {
            programas1.push(row[0]);
          }
        });
        //Llamado a lista dinamica periodo
        var peracad = [];
        data.forEach(row => {
          if (peracad.indexOf(row[2]) == -1) {
            peracad.push(row[2]);
          }
        });

        //Llamado a lista dinamica tipo documento
        var tipodoc = [];
        data.forEach(row => {
          if (tipodoc.indexOf(row[6]) == -1) {
            tipodoc.push(row[6]);
          }
        });
        var template1 = HtmlService.createTemplateFromFile('Calculo_descuento'); //creando una plantilla desde un archivo
        template1.programas1 = programas1;
        template1.peracad = peracad;
        //template1.semdinamico = semdinamico;
        template1.tipodoc = tipodoc;
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;    
      case "15": //formulario de Busqueda Recibos generados
        var template1 = HtmlService.createTemplateFromFile("BuscarRecibos");
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;  
      case "16": //formulario de Busqueda Centros de Costo
        var template1 = HtmlService.createTemplateFromFile("Bus_centro_costos");
        template1.pubUrl = rutaGlobal;
        var output = template1.evaluate();
        break;                           
      default: //Diseño validación de entrada al sitio
        var template6 = HtmlService.createTemplateFromFile("frm_login");
        template6.pubUrl = rutaGlobal + "?p=10";
        var output = template6.evaluate();
        break;
    }

    return output;
  } else {
    const output = HtmlService.createHtmlOutput("<h2>Acceso no permitido</h2><br/><h3>Asegúrate de loggearte con tu cuenta o si crees que es un error contacta al administrador de la aplicación</h3>");
    return output;
  }
}


function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName)
    .getContent()
}

function doPost(e) {

  let params = e.parameter;
  let pantalla = params.p;
  console.log('Que es esto doPost: ' + e.parameter.rutadoc + ' ruta imagen: ' + e.parameter.rutadoc2);

  switch (params.p) {
    case "1":  //Formulario de ingreso de información para las secretarias

      var SS = SpreadsheetApp.getActiveSpreadsheet();
      var sheetRegistro = SS.getSheetByName('Registro Retiros')
      var unsorted_length = sheetRegistro.getLastRow();
      console.log(unsorted_length);
      var id = unsorted_length;
      var fecReg = new Date().toLocaleString();
      var periodoAcad = e.parameter.periodoAcad.toString();
      var progAcademico = e.parameter.progAcademico;
      var fecsolretiro = e.parameter.fecsolretiro.toLocaleString();
      var semestre = e.parameter.semestre;
      var jornada = e.parameter.jorCompletar;
      var identificacion = e.parameter.identificacion;
      var tipoDocumento = e.parameter.tipoDocumento;
      var apellidosNombres = e.parameter.apellidosNombres.toUpperCase();
      var codigoEstudiantil = e.parameter.codigoEstudiantil.toUpperCase();
      var genero = e.parameter.genero;
      var celular = e.parameter.celular;
      var correoElectronico = e.parameter.correoElectronico;
      var obsPrograma = e.parameter.obsPrograma + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" });
      var novedadesNotas = e.parameter.novedadesNotas;
      var consinefectos = ''; //se establece responsabilidad a la vicerrectoria academica
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado';
      var rutadoc = e.parameter.rutadoc;
      //-- Bloqueo por pruebas
      sheetRegistro.appendRow([id, fecReg, periodoAcad, progAcademico, fecsolretiro, semestre, jornada, identificacion, tipoDocumento, apellidosNombres, codigoEstudiantil, genero, celular, correoElectronico, obsPrograma, novedadesNotas, consinefectos, demaspartes, rutadoc]);
      //guardando campos adicionales del formulario
      var custIdsN = sheetRegistro.getRange(2, 1, sheetRegistro.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      var posIndexN = custIdsN.indexOf(id.toString().toLowerCase());
      var rowNumberN = posIndexN === -1 ? 0 : posIndexN + 2;
      console.log('Datos id:' + id + ' rowNumberN' + rowNumberN + '  parametro: ' + e.parameter.rutadoc2);
      sheetRegistro.getRange(rowNumberN, 38, 1, 2).setValues([[
        e.parameter.rutadoc2, e.parameter.plan
      ]]);
      var template = HtmlService.createTemplateFromFile('RegistroTerminadoCC');
      template.pubUrl = rutaGlobal + "?p=3"; //esto se agrega en el post del formulario RegistroTerminado
      template.urlG = rutaGlobal;
      var output = template.evaluate();
      //Recogiendo datos para envio personalizado
      var { origen, destino, emaildestino } = getEnvio(3);
      //------Llamando metodo para envio de correo electrónico
      if (demaspartes == "Aceptado") {
        envioCorreo('1', id, progAcademico, apellidosNombres, codigoEstudiantil, origen, destino, emaildestino);
      } else {
        envioCorreo('7', id, progAcademico, apellidosNombres, codigoEstudiantil, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 
      break;
    case "2": //formulario de informe de datos
      var template1 = HtmlService.createTemplateFromFile("RegistroPA");
      template1.pubUrl = rutaGlobal + "?p=1";
      var output = template1.evaluate();
      break;
    case "3": //formulario de informe Registro en Blanco
      //Llamado a lista dinamica programa
      var data = sheetBD.getDataRange().getValues();
      var programas1 = [];
      data.forEach(row => {
        if (programas1.indexOf(row[0]) == -1) {
          programas1.push(row[0]);
        }
      });
      //Llamado a lista dinamica periodo
      var peracad = [];
      data.forEach(row => {
        if (peracad.indexOf(row[2]) == -1) {
          peracad.push(row[2]);
        }
      });

      //Llamado a lista dinamica semestre
      var semdinamico = [];
      data.forEach(row => {
        if (semdinamico.indexOf(row[3]) == -1) {
          semdinamico.push(row[3]);
        }
      });

      //Llamado a lista dinamica tipo documento
      var tipodoc = [];
      data.forEach(row => {
        if (tipodoc.indexOf(row[6]) == -1) {
          tipodoc.push(row[6]);
        }
      });
      var template1 = HtmlService.createTemplateFromFile("RegistroPA");
      template1.pubUrl = rutaGlobal + "?p=1"; //esto se agrega al formulario de 
      template1.programas1 = programas1;
      template1.peracad = peracad;
      template1.semdinamico = semdinamico;
      template1.tipodoc = tipodoc;
      var output = template1.evaluate();
      break;
    case "4": //formulario de revision credito y cartera
      // Ubicandose en la posicion del registro para realizar las modificaciones
      //console.log('Valor que llega: ' + params.codigoRetiro1 + '- ' + params.progAcademico1 + ' - ' + e.parameter.apellidosNombres1 + ' - ' + e.parameter.codigoEstudiantil1);
      var id = params.codigoRetiro1;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ws = ss.getSheetByName("Registro Retiros");
      const custIds = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      const posIndex = custIds.indexOf(id.toString().toLowerCase());
      const rowNumber = posIndex === -1 ? 0 : posIndex + 2;
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado';
      //Se ubica en la fila a modificar partiendo de la columan 19, para un solo registro, y avanza 4 posiciones hacia adelante
      ws.getRange(rowNumber, 20, 1, 4).setValues([[
        e.parameter.obsCreditoCartera + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }),
        e.parameter.debeMatricula,
        e.parameter.valorPendiente.toString(),
        (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado'
      ]]);
      ws.getRange(rowNumber, 18, 1, 1).setValues([['Completado']]);
      var template1 = HtmlService.createTemplateFromFile("RegistroTerminadoCC");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      //Recogiendo datos para envio personalizado
      var { origen, destino, emaildestino } = getEnvio(4);
      console.log('Email desde codigo: ' + emaildestino);
      //------Llamando metodo para envio de correo electrónico
      //alert('Realizando consulta Origen: '+origen, ' Destino: '+destino+' emaildestino:'+emaildestino);
      if (demaspartes == "Aceptado") {
        // 2,3,4 5,6 : envio normal 7: anulacion envio 6: cierre proceso
        envioCorreo('2', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      } else {
        envioCorreo('7', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 


      break;
    case "5": //formulario Acompañamiento Integral
      // Ubicandose en la posicion del registro para realizar las modificaciones
      //console.log('Valor que llega: ' + params.codigoRetiro1);
      var id = params.codigoRetiro1;
      var progAcademico1 = e.parameter.progAcademico1;
      var apellidosNombres1 = e.parameter.apellidosNombres1.toUpperCase();
      var codigoEstudiantil1 = e.parameter.codigoEstudiantil1;
      var ss1 = SpreadsheetApp.getActiveSpreadsheet();
      var ws1 = ss1.getSheetByName("Registro Retiros");
      var custIds1 = ws1.getRange(2, 1, ws1.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      var posIndex1 = custIds1.indexOf(id.toString().toLowerCase());
      var rowNumber1 = posIndex1 === -1 ? 0 : posIndex1 + 2;
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado';
      //Se ubica en la fila a modificar partiendo de la columna 24, para un solo registro, y avanza 4 posiciones hacia adelante
      ws1.getRange(rowNumber1, 24, 1, 4).setValues([[
        e.parameter.obsAI + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }),
        e.parameter.estadoRetiro,
        (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado',
        e.parameter.acompPrevio

      ]]);
      ws1.getRange(rowNumber1, 23, 1, 1).setValues([['Completado']]);
      var template1 = HtmlService.createTemplateFromFile("RegistroTerminadoCC");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      //Recogiendo datos para envio personalizado
      var { origen, destino, emaildestino } = getEnvio(5);
      //------Llamando metodo para envio de correo electrónico
      if (demaspartes == "Aceptado") {
        envioCorreo('3', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      } else {
        envioCorreo('7', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 
      break;
    case "6": //Formulario Vice. Académica
      // Ubicandose en la posicion del registro para realizar las modificaciones
      //console.log('Valor que llega: ' + params.codigoRetiro1);
      var id = params.codigoRetiro1;
      var progAcademico1 = e.parameter.progAcademico1;
      var apellidosNombres1 = e.parameter.apellidosNombres1.toUpperCase();
      var codigoEstudiantil1 = e.parameter.codigoEstudiantil1;
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var ws2 = ss2.getSheetByName("Registro Retiros");
      var custIds2 = ws2.getRange(2, 1, ws2.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      var posIndex2 = custIds2.indexOf(id.toString().toLowerCase());
      var rowNumber2 = posIndex2 === -1 ? 0 : posIndex2 + 2;
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado';
      //Se ubica en la fila a modificar partiendo de la columan 28, para un solo registro, y avanza 2 posiciones hacia adelante
      ws2.getRange(rowNumber2, 28, 1, 2).setValues([[
        e.parameter.obsVA + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }),
        (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado'
      ]]);
      ws2.getRange(rowNumber2, 26, 1, 1).setValues([['Completado']]);
      var template1 = HtmlService.createTemplateFromFile("RegistroTerminadoCC");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      //Recogiendo datos para envio personalizado
      var { origen, destino, emaildestino } = getEnvio(6);
      //------Llamando metodo para envio de correo electrónico
      if (demaspartes == "Aceptado") {
        envioCorreo('4', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      } else {
        envioCorreo('7', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 
      break;
    case "7": //Formulario Vice. Administrativa y Financiera
      // Ubicandose en la posicion del registro para realizar las modificaciones
      //console.log('Valor que llega: ' + params.codigoRetiro1);
      var id = params.codigoRetiro1;
      var progAcademico1 = e.parameter.progAcademico1;
      var apellidosNombres1 = e.parameter.apellidosNombres1.toUpperCase();
      var codigoEstudiantil1 = e.parameter.codigoEstudiantil1;
      var ss3 = SpreadsheetApp.getActiveSpreadsheet();
      var ws3 = ss3.getSheetByName("Registro Retiros");
      var custIds3 = ws3.getRange(2, 1, ws3.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      var posIndex3 = custIds3.indexOf(id.toString().toLowerCase());
      var rowNumber3 = posIndex3 === -1 ? 0 : posIndex3 + 2;
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado';
      //Se ubica en la fila a modificar partiendo de la columan 19, para un solo registro, y avanza 4 posiciones hacia adelante
      ws3.getRange(rowNumber3, 30, 1, 4).setValues([[
        e.parameter.obsVF + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }),
        e.parameter.casoEsp,
        e.parameter.concepSaldo,
        (e.parameter.demaspartes == 'on') ? 'Aceptado' : 'Rechazado'
      ]]);
      ws3.getRange(rowNumber3, 29, 1, 1).setValues([['Completado']]);
      var template1 = HtmlService.createTemplateFromFile("RegistroTerminadoCC");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      //Recogiendo datos para envio personalizado
      var { origen, destino, emaildestino } = getEnvio(7);
      //------Llamando metodo para envio de correo electrónico
      if (demaspartes == "Aceptado") {
        envioCorreo('5', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      } else {
        envioCorreo('7', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 
      break;
    case "8": //Formulario Registro y Control Académico
      // Ubicandose en la posicion del registro para realizar las modificaciones
      //console.log('Valor que llega: ' + params.codigoRetiro1);
      var id = params.codigoRetiro1;
      var progAcademico1 = e.parameter.progAcademico1;
      var apellidosNombres1 = e.parameter.apellidosNombres1.toUpperCase();
      var codigoEstudiantil1 = e.parameter.codigoEstudiantil1;
      var debeMatricula1 = e.parameter.debeMatricula1;
      var ss4 = SpreadsheetApp.getActiveSpreadsheet();
      var ws4 = ss4.getSheetByName("Registro Retiros");
      var custIds4 = ws4.getRange(2, 1, ws4.getLastRow() - 1, 1).getValues().map(r => r[0].toString().toLowerCase());
      var posIndex4 = custIds4.indexOf(id.toString().toLowerCase());
      var rowNumber4 = posIndex4 === -1 ? 0 : posIndex4 + 2;
      var demaspartes = (e.parameter.demaspartes == 'on') ? 'Cierre' : 'Rechazado';
      //Se ubica en la fila a modificar partiendo de la columan 19, para un solo registro, y avanza 4 posiciones hacia adelante
      ws4.getRange(rowNumber4, 34, 1, 2).setValues([[
        e.parameter.obsRC + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" }),
        (e.parameter.demaspartes == 'on') ? 'Cierre' : 'Rechazado'
      ]]);
      ws4.getRange(rowNumber4, 33, 1, 1).setValues([['Completado']]);
      //Validando si se envia email de cierre o se envia correo de retiro anulado

      var template1 = HtmlService.createTemplateFromFile("RegistroFinalizado");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      //Recogiendo datos para envio personalizado
      //Recuerde la estructura en la hoja Usuarios campo Correo destino va en el orden
      // PA - CC - TS(Tesoreria)
      var { origen, destino, emaildestino } = getEnvio(8);
      //==============================================================================
      // Validando envio a tesoreria o credito y cartera
      var separarMail = emaildestino.split(',');
      // Si presenta credito pendiente, se debe enviar al final al correo de credito y cartera para saber la decision tomada
      if (debeMatricula1.toString().toLowerCase() === "si") {
        var emaildestinoF = separarMail[0] + ',' + separarMail[1];
        destino = destino+'-Oficina de Crédito, Cartera y Cobranzas';
      } else {
        // Si el estudiante no debe nada se va directo a tesoreria
        var emaildestinoF = separarMail[0] + ',' + separarMail[2];
        destino = destino+'-Oficina de Tesorería';
      }
      console.log('Correos que se enviara mensaje: ' + emaildestinoF +' Comparar cadena:'+debeMatricula1.toString().toLowerCase());
    //==============================================================================

      //------Llamando metodo para envio de correo electrónico
      if (demaspartes == "Cierre") {
        envioCorreo('6', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestinoF);
        crearPdf_Retiro(id);
        //crearPdf_RetiroEstudiante(id);
      } else {
        console.log('Preparando camino para anular correo');
        envioCorreo('7', id, params.progAcademico1, params.apellidosNombres1, params.codigoEstudiantil1, origen, destino, emaildestino);
      }
      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado 
      break;
    case "9": //recibiendo datos de notas para ser guardados
      if (typeof e !== 'undefined') {
        Logger.log(e.parameters);
      }
      //var jsonString = e.postData.getDataAsString();
      //var jsonData = JSON.parse(jsonString);
      //sheet.appendRow(['Data1:', jsonData.Data1]); // Just an example
      //console.log('Valor que llega: ' + jsonData);

      //var template1 = HtmlService.createTemplateFromFile("RegistroFinalizado");
      template1.pubUrl = rutaGlobal + "?p=3"; //esto se agrega al formulario de Registro Terminado
      template1.urlG = rutaGlobal;
      var output = template1.evaluate();
      break;
    case "10": //Diseño pagina de aterrizaje para el sitio
      // console.log('Valor que llega datEmail: ' + e.parameter.datEmail + ': ' + Session.getActiveUser().getEmail());
      var correoUsuario = e.parameter.datEmail.toString().trim();
      var email = Session.getActiveUser().getEmail();
      if (email == correoUsuario) {
        let registrosbusqueda = [];
        const SS2 = SpreadsheetApp.getActiveSpreadsheet();
        const sheetRegistros = SS2.getSheetByName("Usuarios");
        //const usuarioBD = sheetRegistros.getDataRange().getDisplayValues();
        var headerRowNumber = 1;
        //How do I exclude a header row from getDataRange()?
        const usuarioBD = sheetRegistros.getDataRange().offset(headerRowNumber, 0, sheetRegistros.getLastRow() - headerRowNumber).getValues();

        usuarioBD.forEach(usuarioBD => {
          //console.log('Comparacion:' + usuarioBD[0].toString() + ' --- ' + correoUsuario.toString());
          if (usuarioBD[0].toString() === correoUsuario.toString()) {
            registrosbusqueda.push(usuarioBD);
          }
        })
        //console.log('Correo 1 : ' + registrosbusqueda);
        //console.log(' \n email 1:' + email);
        //console.log('\n activo: ' + registrosbusqueda[2]);
        console.log('Comparacion: registrosbusqueda[0][0].toString():' + registrosbusqueda[0][0].toString() + '\n correoUsuario.toString(): ' + correoUsuario.toString() + '\n registrosbusqueda[0][0].toString():' + registrosbusqueda[0][0].toString() + ' email.toString():' + email.toString() + '\n registrosbusqueda[0][3]:' + registrosbusqueda[0][3].toString().toUpperCase());
        if ((registrosbusqueda[0][0].toString() == correoUsuario.toString()) && (registrosbusqueda[0][0].toString() == email.toString()) && (registrosbusqueda[0][3].toString().toUpperCase() == 'TRUE')) {
          console.log('\n Es igual');
          //-----------------------------------------------
          var data = registrosbusqueda[0][1];
          Logger.log('Datos Usuario unico: -->' + registrosbusqueda[0][2].toString());
          //Estableciendo en variable constante la forma de mostrar los programas a trabajar
          const objprog = JSON.parse(data);
          switch (registrosbusqueda[0][2].toString()) {
            case '1':
              console.log('Administrador');
              //En el caso de ser administrador
              //var sitiohmt = "BuscarRetirosAdm";
              var templateF = HtmlService.createTemplateFromFile("BuscarRetirosAdm");
              templateF.pubUrl = rutaGlobal;
              templateF.programaBD = objprog;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              break;
            case '2':
              console.log('Secretarias');
              //En el caso de ser secretari@ de programa
              //var sitiohmt = "BuscarRetiros";
              var templateF = HtmlService.createTemplateFromFile("BuscarRetiros");
              templateF.pubUrl = rutaGlobal;
              templateF.programaBD = objprog;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              //var rutaApp = rutaGlobal; // + "?p=3";
              break;
            case '3':
              console.log('otras dependencias');
              //En el caso de ser las demas oficinas verificadoras
              //var sitiohmt = "EstadosRetiros";
              templateF = HtmlService.createTemplateFromFile("EstadosRetiros");
              templateF.pubUrl = rutaGlobal;
              //templateF.programaBD = objprog;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              //var rutaApp = rutaGlobal;// + "?p=11";
              break;
            case '4': // Acceso a Guia telefónica otras dependencias
              //En el caso de ser las demas oficinas verificadoras
              //var sitiohmt = "EstadosRetiros";
              templateF = HtmlService.createTemplateFromFile("Bus_telefono");
              templateF.pubUrl = rutaGlobal;
              //templateF.programaBD = objprog;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              //var rutaApp = rutaGlobal;// + "?p=11";
              break;
            case '5': // Solo Tarifas
              //En el caso de ser las demas oficinas verificadoras
              //var sitiohmt = "EstadosRetiros";
              templateF = HtmlService.createTemplateFromFile("Bus_tarifas_matricula");
              templateF.pubUrl = rutaGlobal;
              //templateF.programaBD = objprog;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              //var rutaApp = rutaGlobal;// + "?p=11";
              break;              
            default:
              console.log('LOS DEMAS');
              templateF = HtmlService.createTemplateFromFile("Bus_telefono");
              templateF.pubUrl = rutaGlobal;
              var output = templateF.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
              break;
          }//fin switch
          //-----------------------------------------------

        } //fin acceso correcto

      }
      else {
        console.log('\n NO Es igual');
        var template2 = HtmlService.createTemplateFromFile("UsuarioIncorrecto");
        template2.pubUrl = rutaGlobal; //se define la url para que al hacer clic vuelva a intentar logearse con una contraseña correcta
        var output = template2.evaluate();
      }
      break;
      case "14": //guardado de datos para generar recibo de pago
      var SS = SpreadsheetApp.getActiveSpreadsheet();
      var sheetRegistro = SS.getSheetByName('GenRec')
      var unsorted_length = sheetRegistro.getLastRow();
      console.log(unsorted_length);
      var id = unsorted_length;
      var fecReg = new Date().toLocaleString();
      var periodoAcad = "2023-1";//e.parameter.periodoAcad.toString();
      var progAcademico = e.parameter.progAcademico;
      var semestre = e.parameter.semestre;
      var jornada = e.parameter.jorCompletar;
      var identificacion = e.parameter.identificacion;
      var apellidosNombres = e.parameter.apellidosNombres.toUpperCase();
      var codigoEstudiantil = e.parameter.codigoEstudiantil.toUpperCase();
      var genero = e.parameter.genero;
      var tipodat = e.parameter.tipodat;
      var sem_cursar = e.parameter.semestre_cal;
      var creditos_cal = e.parameter.creditos_cal;
      // Armando estructura de valores para generación de recibo
      var val_5 = e.parameter.des_5_env;
      var val_3 = e.parameter.des_3_env;
      var val_2 = e.parameter.des_2_env;
      var val_ord = e.parameter.des_ord_env;
      var url_5 = "https://barcode.tec-it.com/barcode.ashx?data=(415)7709998636576(8020)"+zfill(id, 6)+"(8020)"+zfill(identificacion, 12)+"(3900)"+zfill(val_5, 10)+"(96)20221215&code=EANUCC128&translate-esc=true&unit=Min&imagetype=Png&modulewidth=1";
      var codbarras_5 = "=IMAGE(\""+url_5+"\")";
      var url_3 = "https://barcode.tec-it.com/barcode.ashx?data=(415)7709998636576(8020)"+zfill(id, 6)+"(8020)"+zfill(identificacion, 12)+"(3900)"+zfill(val_3, 10)+"(96)20221230&code=EANUCC128&translate-esc=true&unit=Min&imagetype=Png&modulewidth=1";
      var codbarras_3 = "=IMAGE(\""+url_3+"\")";  
      var url_2 = "https://barcode.tec-it.com/barcode.ashx?data=(415)7709998636576(8020)"+zfill(id, 6)+"(8020)"+zfill(identificacion, 12)+"(3900)"+zfill(val_2, 10)+"(96)20230112&code=EANUCC128&translate-esc=true&unit=Min&imagetype=Png&modulewidth=1";
      var codbarras_2 = "=IMAGE(\""+url_2+"\")";        
      var obsPrograma = e.parameter.obsPrograma + '\n Registro: ' + new Date().toLocaleDateString('es-CO', { weekday: "long", year: "numeric", month: "short", day: "numeric" });

       var correoElectronico = e.parameter.correoElectronico; //dato para enviar información al estudiante
       // Guardando datos en hoja 
      sheetRegistro.appendRow([id, fecReg, periodoAcad, progAcademico, semestre, jornada, identificacion, apellidosNombres, codigoEstudiantil, genero, tipodat, val_5, val_3, val_2, url_5, codbarras_5, url_3, codbarras_3, url_2, codbarras_2, obsPrograma, sem_cursar, creditos_cal]);
      //Fin de guardado
      var template = HtmlService.createTemplateFromFile('RespuestaRecAnt'); 
      template.pubUrl = rutaGlobal + "?p=15"; //esto se agrega en el post del formulario RegistroTerminado PENDIENTE
      template.urlG = rutaGlobal;
      template.codigo = codigoEstudiantil;
      template.apellidosNombres = apellidosNombres;
      template.correoElectronico = correoElectronico;
      var output = template.evaluate();
      //Recogiendo datos para envio personalizado
      var origenC = Session.getActiveUser().getEmail();
      //------Llamando metodo para envio de correo electrónico

       // envioCorreo('Rec', id, progAcademico, apellidosNombres, codigoEstudiantil, origenC, '', correoElectronico);
        crearPdf_Pago_Anticipado(id, val_ord); //REVISAR PLANTILLA

      //-- En la funcion de envio de correo se envia un numero que indica a que dependencia se va el correo y el id del formulario que se debe editar o revisar
      //------Fin de llamado       
      break;
    default: {
      Logger.log("Entro por Deafault");
      var template3 = HtmlService.createTemplateFromFile("frm_login");
      template3.pubUrl = rutaGlobal;
      var output = template3.evaluate();
    }
  }

  return output;
}

function showDialog() {                                                     //for show the msg to the user
  var template = HtmlService.createTemplateFromFile("Dialog").evaluate();   //file html
  SpreadsheetApp.getUi().showModalDialog(template, "Cargar Documento");           //Show to user, add title
}

// Funcion para cargar archivo firmado por el estudiante
function uploadFilesToGoogleDrive(data, name, type) {                          //function to call on front side
  var datafile = Utilities.base64Decode(data)                               //decode data from Base64
  var blob2 = Utilities.newBlob(datafile, type, name);                      //create a new blob with decode data, name, type
  var folder = DriveApp.getFolderById("14EezGw7VU2giA8xasJgmZOt0tiFoC-Ih"); //Get folder of destiny for file (final user need access before execution)

  //console.log(' identificacion: ' + ideusu);
  // var fileName1 = ideusu;// Identificacion
  // var fileName2 = ideusu // Periodo

  //var newFile = folder.createFile(blob2);  //proximo a cambiar
  // var newFile = folder.createFile(blob2).setName(`${fileName1} _ ${fileName2} _Retiro`); //estableciendo nombre randomico al documento
  var newFile = folder.createFile(blob2).setName(`prueba_Retiro`); //estableciendo nombre randomico al documento

  return newFile.getUrl()                                                    //Return URL
}


//Ceros a la izquierda manteniendo formato
function zfill(number, width) {
  var numberOutput = Math.abs(number); /* Valor absoluto del número */
  var length = number.toString().length; /* Largo del número */
  var zero = "0"; /* String de cero */

  if (width <= length) {
    if (number < 0) {
      return ("-" + numberOutput.toString());
    } else {
      return numberOutput.toString();
    }
  } else {
    if (number < 0) {
      return ("-" + (zero.repeat(width - length)) + numberOutput.toString());
    } else {
      return ((zero.repeat(width - length)) + numberOutput.toString());
    }
  }
}

function prueba1() {
  var idBus = 3;
  var email = Session.getActiveUser().getEmail();
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("Usuarios");
  //const usuarioBD = sheetRegistros.getDataRange().getDisplayValues();
  var headerRowNumber = 1;
  const usuarioBD = sheetRegistros.getDataRange().offset(headerRowNumber, 0, sheetRegistros.getLastRow() - headerRowNumber).getValues();

  //validando a que solo entre encuentre el valor y salga
  var BreakException = {};
  try {
    usuarioBD.forEach(usuarioBD => {
      //console.log('Comparacion:' + usuarioBD[0].toString() + ' --- ' + email);
      if ((usuarioBD[0].toString() === email.toString()) && (usuarioBD[5] === idBus)) {
        registrosbusqueda.push(usuarioBD);
        throw BreakException;
      }
    });
  } catch (e) {
    if (e !== BreakException) throw e;
  }
  //console.log('Correo 1 : ' + registrosbusqueda[0][0]);
  //console.log(' \n email 1:' + email);
  //console.log('\n activo: ' + registrosbusqueda[0][3]);
  var data1 = registrosbusqueda[0][6];//area origen
  var data2 = registrosbusqueda[0][7];//area destino  
  var data3 = registrosbusqueda[0][8];//correo destino

  Logger.log('Datos Usuario unico: -->' + data1 + '\n Destino:' + data2 + '\n Correo Destino:' + data3);

}

 function remove_accents(strAccents) {
    var strAccents = strAccents.split('');
    var strAccentsOut = new Array();
    var strAccentsLen = strAccents.length;
    var accents =    "ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëÇçðÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž";
    var accentsOut = "AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeCcdDIIIIiiiiUUUUuuuuNnSsYyyZz";
    for (var y = 0; y < strAccentsLen; y++) {
        if (accents.indexOf(strAccents[y]) != -1) {
            strAccentsOut[y] = accentsOut.substr(accents.indexOf(strAccents[y]), 1);
        } else
            strAccentsOut[y] = strAccents[y];
    }
    strAccentsOut = strAccentsOut.join('');

    return strAccentsOut;
}


