function doGet(){

  var template = HtmlService.createTemplateFromFile('index');
  template.data = getSheetData();
  var output = template.evaluate();
  return output
  
}

function include(fileName){
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function getSheetData(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPacientes = ss.getSheetByName('Activos');
  var dataPacientes = hojaPacientes.getDataRange().getValues();
  dataPacientes.shift();

  return dataPacientes;

}

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('Remisiones Foscal')
    .addItem('Importar correos nuevos', 'getCorreos')
    .addItem('Enviar respuestas pendientes', 'sendRespuestas')
    .addItem('Enviar respuestas REPLY', 'responderCorreos')
    .addItem('Acerca de', 'acercaDe')

    .addToUi();

}


function getCorreos() {
  //Consulta la bandeja de entrada del correo y descarga aquellos mensajes marcados con "Recibidos"

  SpreadsheetApp.getActiveSpreadsheet().toast('Importando correos...', 'Aviso');

  var labelRecibidos = GmailApp.getUserLabelByName("@Recibidos");
  var labelDescargados = GmailApp.getUserLabelByName("@Descargados");  //Inactivo porque está comentada la línea 36

  var threads = labelRecibidos.getThreads();

  for (var i = threads.length - 1; i >= 0; i--) {

    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {

      var message = messages[j];
      extractDetails(message);

    }

    //Inactivos en pruebas.  Quitar el comment para PRD ***
    //threads[i].removeLabel(labelRecibidos);
    //labelDescargados.addToThread(threads[i]);

  }

  SpreadsheetApp.getActiveSpreadsheet().toast('¡Importación terminada!', '* Hecho *');

}


function extractDetails(message) {
  //Extrae los detalles de un mensaje recibido como parámetro y los graba en la sábana

  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();
  var bodyContents = message.getPlainBody();
  var idMensaje = message.getId();
  var estado = "Nuevo";

  var folderPaciente = getAdjuntos(message);
  var urlFolder = folderPaciente.getUrl();
  var email = extraerCorreo(message);

  var hojaActivos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activos");

  var ultimaFila = hojaActivos.getLastRow() + 1;
  var formulaDuplicados = '=IF(H' + ultimaFila + '="";"";IF(COUNTIF(H:H;H' + ultimaFila +')>1;"REPETIDA";"NUEVA"))';

  hojaActivos.appendRow([dateTime, senderDetails, email, idMensaje, subjectText, bodyContents, urlFolder,,,,,,formulaDuplicados,,,estado]);

  

  //Inactivo para pruebas.  Quitar comment para PRD ***
  //GmailApp.markMessageRead(message);

}


function getAdjuntos(message) {
  //Descarga los adjuntos al mensaje de correo que recibe como parámetro y los guarda en Drive en una carpeta creada con la función crearFolder()

  Utilities.sleep(1000);
  var folderPaciente = crearFolder();

  var attachments = message.getAttachments();


  for (var i = 0; i < attachments.length; i++) {

    var attachment = attachments[i];

    var attachmentBlob = attachment.copyBlob();
    var file = DriveApp.createFile(attachmentBlob);

    folderPaciente.addFile(file);

    var j = i + 1;
    SpreadsheetApp.getActiveSpreadsheet().toast('Descargando adjunto ' + j + ' de ' + attachments.length, 'Aviso');

  }

  return folderPaciente;

}


function crearFolder() {
  //Crea folder nuevo en la carpeta de la App, identificado con fecha y hora actuales 

  var fechaHoy = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy - HH:mm:ss.SS");
  var folderApp = DriveApp.getFolderById('1W-9NxEtof6hTo8-V4mrttbtuGnDCOeCM');
  folderNuevo = folderApp.createFolder(fechaHoy);
  var folderId = folderNuevo.getId();

  return folderNuevo;

}

function extraerCorreo(mensaje) {
  //Aísla la dirección email de la celda especificada en "data[i][1]"

  var remitente = mensaje.getFrom();
  var correo = remitente.match(/\S+@\S+\.\S+/g);
  var email = correo[0];
  email = email.replace("<", "");
  email = email.replace(">", "");
  return email;

}

function acercaDe() {
  //Créditos 

  SpreadsheetApp.getActiveSpreadsheet().toast('Por: División de Inteligencia Empresarial - Observatorio de Salud Pública de Santander', 'REMISIONES FOSCAL');

}

function sendRespuestas() {
  //Envía las respuestas de los casos que esten marcadoos como "Listo" y los marca como "Cerrado"

  var libroRemisiones = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActivos = libroRemisiones.getSheetByName("Activos");

  libroRemisiones.toast('Enviando respuestas...', 'Aviso');

  var data = hojaActivos.getDataRange().getValues();

  for (i = 1; i <= data.length - 1; i++) {

    var estado = data[i][15];

    if (estado == "Listo") {

      email = data[i][2];
      respuesta = data[i][16];
      observ = data[i][18];
      paciente = data[i][9];

      GmailApp.sendEmail(email, "Respuesta a solicitud paciente: " + paciente, "Cordial saludo.\r \r La respuesta a la solicitud es: " + respuesta + ". \r Observaciones: " + observ + ".  \r \r Gracias por contactar a la Clínica Foscal");

      var celdaEstado = hojaActivos.getRange(i + 1, 16);

      celdaEstado.setValue("Cerrado");

      var hojaCerrados = libroRemisiones.getSheetByName("Cerrados");

      hojaCerrados.appendRow([data[i][0],data[i][1],data[i][2],data[i][3],data[i][4],data[i][5],data[i][6],data[i][7],data[i][8],data[i][9],data[i][10],data[i][11],data[i][12],data[i][13],data[i][14],"Cerrado",data[i][16],data[i][17],data[i][18], data[i][19],data[i][20],data[i][21]]);
      
      hojaActivos.deleteRow(i+1);


    }

  }

  libroRemisiones.toast('¡Respuestas enviadas!', '* Hecho *');

}

function responderCorreos() {
  //Envía respuesta a la solicitud como "Reply" al correo original

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet();

  ss.toast('Enviando respuestas...', 'Aviso');

  var data = ss.getDataRange().getValues();

  for (i = 1; i <= data.length - 1; i++) {

    var estado = data[i][6];

    if (estado == "Listo") {

      email = data[i][2];
      respuesta = data[i][16];
      observ = data[i][18];
      idMensaje = data[i][3];

      GmailApp.getMessageById(idMensaje).reply("Cordial saludo.\r \r La respuesta a la solicitud es: " + respuesta + ". \r El motivo de la respuesta es: " + observ + ".  \r \r Gracias por contactar a la Clínica Foscal");

      var celdaEstado = hoja.getRange(i + 1, 7);

      celdaEstado.setValue("Cerrado");

    }

  }

  ss.toast('¡Respuestas enviadas!', '* Hecho *');

}
