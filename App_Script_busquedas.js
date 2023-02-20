
//Buscando registros para revisar y agregar información adicional
function buscarRegistro(codigo = "7212587") {
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("Registro Retiros");
  const valoresretiro = sheetRegistros.getDataRange().getDisplayValues();
  // console.log(valoresretiro);

  valoresretiro.forEach(valoresretiro => {
    if (valoresretiro[10] === codigo) {
      registrosbusqueda.push(valoresretiro);
    }
  })
  console.log(registrosbusqueda);
  return registrosbusqueda;
}

function buscarRecibo(codigo = "7212587") {
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("GenRec");
  const valoresretiro = sheetRegistros.getDataRange().getDisplayValues();
  // console.log(valoresretiro);

  valoresretiro.forEach(valoresretiro => {
    if (valoresretiro[8] === codigo) {
      registrosbusqueda.push(valoresretiro);
    }
  })
  console.log(registrosbusqueda);
  return registrosbusqueda;
}

//Buscando oficinas para contacto institucional
function getData() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMantenimientos = SS.getSheetByName('Guia');
  const dataMantenimientos = sheetMantenimientos.getDataRange().getDisplayValues();
  dataMantenimientos.shift();
  return dataMantenimientos;
}

//Buscando oficinas para contacto institucional
function getCostos() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMantenimientos = SS.getSheetByName('Centro_Costos');
  const dataMantenimientos = sheetMantenimientos.getDataRange().getDisplayValues();
  dataMantenimientos.shift();
  return dataMantenimientos;
}
//Buscando Tarifas de matricula
function getDataMat() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMantenimientos = SS.getSheetByName('VLR.Matriculas');
  const dataMantenimientos = sheetMantenimientos.getDataRange().getDisplayValues();
  dataMantenimientos.shift();
  return dataMantenimientos;
}

const getEnvio = (idBus) => {
  //var idBus=3;
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

  //Logger.log('Datos Usuario unico: -->' + data1+'\n Destino:'+data2+'\n Correo Destino:'+data3);
  return {
    origen: data1,
    destino: data2,
    emaildestino: data3
  };
}

//Buscando codigo para completar formulario de solicitud de retiro
function buscarCodigo(codigo = "6183116") {
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("Matriculas");
  const valoresretiro = sheetRegistros.getDataRange().getDisplayValues();
  //console.log(valoresretiro);

  valoresretiro.forEach(valoresretiro => {
    if (valoresretiro[0] === codigo) {
      registrosbusqueda.push(valoresretiro);
    }
  })
  console.log(registrosbusqueda);
  return registrosbusqueda;
}

//Buscando programa para iniciar calculo de valores de matricula
function buscarPrograma(programa = "ESPECIALIZACIÓN EN INFANCIA, CULTURA Y DESARROLLO-2-1-2023") {
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("VLR.Matriculas");
  const valoresretiro = sheetRegistros.getDataRange().getDisplayValues();
  //console.log(valoresretiro);

  valoresretiro.forEach(valoresretiro => {
    if (valoresretiro[0] === programa) {
      registrosbusqueda.push(valoresretiro);
    }
  })
  console.log(registrosbusqueda);
  return registrosbusqueda;
}

//Buscando programa para consultar creditos por semestre
function buscarProgramaCred(programa = "PSICOLOGÍA") {
  let registrosbusqueda = [];
  const SS2 = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistros = SS2.getSheetByName("AsigCreditos");
  const valoresretiro = sheetRegistros.getDataRange().getDisplayValues();
  //console.log(valoresretiro);

  valoresretiro.forEach(valoresretiro => {
    if (valoresretiro[0] === programa) {
      registrosbusqueda.push(valoresretiro);
    }
  })
  console.log(registrosbusqueda);
  return registrosbusqueda;
}