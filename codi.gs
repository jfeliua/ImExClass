function onInstall(e) {
  onOpen(e)
};

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  switch(Session.getActiveUserLocale()){
    case "ca":
      var menu= 'Obrir ImExClass';
      break;
    case "es":
      var menu= 'Abrir ImExClass';
      break;
    default:
      var menu= 'Open ImExClass';
  }  
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem(menu,'barra')  
  .addToUi()
};

function barra(){
 switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='index_ca';
      break;
    case "es":
      var nom_html='index_es';
      break;
    default:
      var nom_html='index_en';
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
  
   SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
};


function tasca(formObject) {
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
  var colmail = formObject.colmail;
  var colgrade = formObject.colgrade;
  var cursid = formObject.combo_curs;
  var imex = formObject.imex;
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='Cal triar un curs de Classroom';
      var curs_m='Curs';
      break;
    case "es":
      var nom_html='Es necesario elegir un curso de Classroom';
      var curs_m='Curso';
      break;
    default:
      var nom_html='It is necessary to choose a Classroom course';
      var curs_m='Course';
  };  
  if (cursid == 0){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    barra();
  }else{
    var buit=Classroom.Courses.Students.list(cursid);
    if (isEmpty(buit)){
      switch(Session.getActiveUserLocale()){
        case "ca":
          var nom_html='Cal triar un curs de Classroom amb alumnes inscrits';
          var curs_m='Curs';
          break;
        case "es":
          var nom_html='Es necesario elegir un curso de Classroom con alumnos inscritos';
          var curs_m='Curso';
          break;
        default:
          var nom_html='It is necessary to choose a Classroom course with registered students';
          var curs_m='Course';
      };
      var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
      barra();
    }else{
      var filgrade = formObject.filgrade;
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('cursid', cursid);
      documentProperties.setProperty('colmail', colmail);
      documentProperties.setProperty('colgrade', colgrade);
      documentProperties.setProperty('filgrade', filgrade);
      documentProperties.setProperty('imex', imex); 
      if (imex=="0"){     
        switch(Session.getActiveUserLocale()){
          case "ca":
            var nom_html='tasca_ca';
            break;
          case "es":
            var nom_html='tasca_es';
            break;
          default:
            var nom_html='tasca_en';
        }; 
      }else{
        switch(Session.getActiveUserLocale()){
          case "ca":
            var nom_html='import_ca';
            break;
          case "es":
            var nom_html='import_es';
            break;
          default:
            var nom_html='import_en';
        }; 
      };     
      var html = HtmlService
      .createTemplateFromFile(nom_html)
      .evaluate();      
      SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass'); 
    };
  };
};

function tasca2() {
  var properties = PropertiesService.getDocumentProperties();
  var imex = properties.getProperty('imex');
  if (imex=="0"){     
    switch(Session.getActiveUserLocale()){
      case "ca":
        var nom_html='tasca_ca';
        break;
      case "es":
        var nom_html='tasca_es';
        break;
      default:
        var nom_html='tasca_en';
    }; 
  }else{
    switch(Session.getActiveUserLocale()){
      case "ca":
        var nom_html='import_ca';
        break;
      case "es":
        var nom_html='import_es';
        break;
      default:
        var nom_html='import_en';
    }; 
  };
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
    
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass'); 
};
  
function accio(formObject) {
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
  var titol = formObject.titol;
  var descripcio = formObject.descripcio;
  var pmax = formObject.pmax;
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('titol', titol);
  documentProperties.setProperty('descripcio', descripcio);
  documentProperties.setProperty('pmax', pmax);
  documentProperties.setProperty('tascaid', "");  //guardo que no hem creat la tasca
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='Cal indicar un nom per la tasca';
      var curs_m='Tasca';
      break;
    case "es":
      var nom_html='Es necesario indicar un nombre para la tarea';
      var curs_m='Tarea';
      break;
    default:
      var nom_html='You need to enter a name for the assigment';
      var curs_m='Assigment';
  }; 
  if (titol == ""){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    tasca2();
  }else{
     exportar();    
  }; 
};

function accio_i(formObject) {
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
  var tascaid_im = formObject.combo_tasca;
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('tascaid_im', tascaid_im);  
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='Cal seleccionar una tasca';
      var curs_m='Tasca';
      break;
    case "es":
      var nom_html='Es necesario seleccionar una tarea';
      var curs_m='Tarea';
      break;
    default:
      var nom_html='You need to select one assigment';
      var curs_m='Assigment';
  }; 
  if (tascaid_im == ""){
    var msg=Browser.msgBox(curs_m,nom_html, Browser.Buttons.OK);
    tasca2();
  }else{
     importar();    
  }; 
};

function exportar(){
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();     
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
  var properties = PropertiesService.getDocumentProperties();
  var cursid = properties.getProperty('cursid');
  var titol = properties.getProperty('titol');
  var tascaid = properties.getProperty('tascaid');
  var descripcio= properties.getProperty('descripcio');
  var pmax = properties.getProperty('pmax');
  var colmail = properties.getProperty('colmail');
  var colgrade = properties.getProperty('colgrade');
  var filgrade = properties.getProperty('filgrade');
  var rang_full = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange()
  var dades_full= rang_full.getValues();
  //si no existeix, creo la tasca
  if (tascaid==""){
    var creo_tasca = {
      "courseId": cursid,
      "title": titol,
      "description": descripcio,
      "maxPoints": pmax,
      "workType":"ASSIGNMENT",
      "state": "PUBLISHED"
    }
    var tasca_creada=Classroom.Courses.CourseWork.create(creo_tasca, cursid); //Només es poden canviar notes de tasques creades per la API
    var tascaid=tasca_creada.id;
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('tascaid', tascaid);    
  }
  var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tascaid); //recuperem totes les tasques de tots els alumnes
  var i=filgrade-1;
  var h=1;
  var totalum=rang_full.getNumRows()-filgrade+1;
  properties.setProperty('totalum', totalum);
  for (var i=filgrade-1; i<rang_full.getNumRows();i++){ //agafo alumne per alumne. A partir del mail, trobarem l'userid. A partir de l'userid,el submissionid
    var nalum=h;
    properties.setProperty('nalum', nalum);
    var html = HtmlService
    .createTemplateFromFile('updating2')
    .evaluate();
    SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
    var mail_st = dades_full[i][colmail-1];
    var nota_st = dades_full[i][colgrade-1];
    if (!(isNaN(nota_st))){
      var pagina=null;
      var ki=0;
      var alumnes = []; 
      do {
        alumnes[ki]=Classroom.Courses.Students.list(cursid,{pageToken:pagina});  //Classroom treu els alumnes de 30 en 30. Cal llegir 30 i després canviar el token per llegir-ne 30 més
        pagina=alumnes[ki].nextPageToken;
        ki++;
      }while (pagina);
      var mail1 = [];
      var userid = [];
      var comptador=0;
      for (var f=0;f<alumnes.length;f++){
        for (var l=0;l<alumnes[f].students.length;l++){
          mail1[comptador]=alumnes[f].students[l].profile.emailAddress;
          userid[comptador]=alumnes[f].students[l].userId;
          comptador++;
        };
      };      
      //var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
      for (var j=0;j<mail1.length;j++){ 
        //agafem alumne per alumne i comparem amb el de la cel·la
        if (mail1[j]==mail_st){
          var userid=userid[j]; //Trobem userid de l'usuari de la cel·la
          for (var k=0; k<tasques_env.studentSubmissions.length;k++){
            var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
            var env_id=tasques_env.studentSubmissions[k].id;
            if (env_us==userid){
              var reso = {'draftGrade':nota_st};
              var extra={'updateMask':'draftGrade'};
              var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tascaid, env_id,extra); // Actualitzem la nota esborrany
              var reso = {'assignedGrade':nota_st};
              var extra={'updateMask':'assignedGrade'};
              var log_class=Classroom.Courses.CourseWork.StudentSubmissions.patch(reso, cursid, tascaid, env_id,extra);// Actualitzaem la nota que veu l'alumne
              var reso = {'return':1};
            }
          }        
        }
      }
      h++;
    }
  }
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='exportat_ca';
      break;
    case "es":
      var nom_html='exportat_es';
      break;
    default:
      var nom_html='exportat_en';
  };  
  var html = HtmlService
    .createTemplateFromFile(nom_html)
    .evaluate();
    
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');   
}

function importar(){
  var html = HtmlService
  .createTemplateFromFile('updating')
  .evaluate();      
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
  var properties = PropertiesService.getDocumentProperties();
  var cursid = properties.getProperty('cursid');
  var tascaid_im = properties.getProperty('tascaid_im');
  var colmail = properties.getProperty('colmail');
  var colgrade = properties.getProperty('colgrade');
  var filgrade = properties.getProperty('filgrade');
  var rang_full = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange()
  var dades_full= rang_full.getValues();
  var tasques_env=Classroom.Courses.CourseWork.StudentSubmissions.list(cursid, tascaid_im); //recuperem totes les tasques de tots els alumnes
  var i=filgrade-1;
  var totalum=rang_full.getNumRows()-filgrade+1;
  properties.setProperty('totalum', totalum);
  for (var i=filgrade-1; i<rang_full.getNumRows();i++){ //agafo alumne per alumne. A partir del mail, trobarem l'userid. A partir de l'userid,el submissionid
    var mail_st = dades_full[i][colmail-1];
    var pagina=null;
    var ki=0;
    var alumnes = []; 
    do {
      alumnes[ki]=Classroom.Courses.Students.list(cursid,{pageToken:pagina});  //Classroom treu els alumnes de 30 en 30. Cal llegir 30 i després canviar el token per llegir-ne 30 més
      pagina=alumnes[ki].nextPageToken;
      ki++;
    }while (pagina);
    var mail1 = [];
    var userid = [];
    var comptador=0;
    for (var f=0;f<alumnes.length;f++){
      for (var l=0;l<alumnes[f].students.length;l++){
        mail1[comptador]=alumnes[f].students[l].profile.emailAddress;
        userid[comptador]=alumnes[f].students[l].userId;
        comptador++;
      };
    };
    //var llista_st = Classroom.Courses.Students.list(cursid); //Agafo la llista d'alumnes
    for (var j=0;j<mail1.length;j++){
       //agafem alumne per alumne i comparem amb el de la cel·la
      if (mail1[j]==mail_st){
        var userid=userid[j]; //Trobem userid de l'usuari de la cel·la
        for (var k=0; k<tasques_env.studentSubmissions.length;k++){
          var env_us=tasques_env.studentSubmissions[k].userId; //busco la tasca de l'usuari de la cel·la
          var env_id=tasques_env.studentSubmissions[k].id;
          if (env_us==userid){
            //Copiem les notes
            var nota = tasques_env.studentSubmissions[k].assignedGrade;
            if (nota!=undefined){
              SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(i+1,colgrade).setValue(nota);    
            };
            var nalum= i;
            properties.setProperty('nalum', nalum);
            html = HtmlService
            .createTemplateFromFile('updating2')
            .evaluate();
            SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass');
          }
        }        
      }
    }
  }
  switch(Session.getActiveUserLocale()){
    case "ca":
      var nom_html='importat_ca';
      break;
    case "es":
      var nom_html='importat_es';
      break;
    default:
      var nom_html='importat_en';
  };  
  var html = HtmlService
  .createTemplateFromFile(nom_html)
  .evaluate();
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'ImExClass'); 
};


function isEmpty(myObject) {
    for(var key in myObject) {
        if (myObject.hasOwnProperty(key)) {
            return false;
        }
    }

    return true;
}
  
