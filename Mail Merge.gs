function myFunction() {
  
  var docTemplateId = "1kmr-26VPKS7JXsvybVU50YRhH0Y9zzdGq1Me51EsPNE"
  var docFinalId = "1bkKOkeoCyGvn9Z_yO5i19GVkoTI4nmJw9fKlC-og130"
  var wsId = "1GjisxWYNefoB1huG7fhUU537BnNuYpEPHxF50Hp21X4"

  var doctTemplate = DocumentApp.openById(docTemplateId);
  var docFinal = DocumentApp.openById(docFinalId); 
  var ws = SpreadsheetApp.openById(wsId).getSheetByName("data");

  var data = wsId.getRange(2,1,ws.getLastRow()-1,7).getValues(); 

  var templateParagraphs = doctTemplate.getBody().getParagraphs();
  docTemplate.getBody().getParagraphs()[0].getType()

  docFinal.getBody().clear();

  data.forEach(function(r){
    creatMailMerge(r[0],r[1],r[6],templateParagraphs,docFinal);
  });


  }


function creatMailMerge(Name,Address, templateParagraphs,docFinal){
  templateParagraphs,forEach(function(p){
    var elType = p.getType(); 

    if(elType == "PARAGRAPH"){
    docFinal.getBody().appendParagraph(
      p
      .copy()
      .replaceText("{Name}",Name)
      .replaceText("{Address}",Address)
      );
  }  else if(elType == "LIST_ITEM"){
  docFinal.getBody().appendParagraphs(
      p
      .copy()
      .replaceText("{Name}",Name)
      .replaceText("{Address}",Address)
      ).setGylphType(DocumentApp.GlyphType.BULLET);
  }





  });

  docFinal.getBody().appendPageBreak();
}
 