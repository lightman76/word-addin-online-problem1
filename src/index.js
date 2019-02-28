/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
  return Word.run(async context => {
   /**
    * Insert your Word code here
    */
    var doc = context.document;
    var endOfBodyRange = doc.body.getRange(Word.RangeLocation.end);
    endOfBodyRange.insertBreak(Word.BreakType.page, Word.InsertLocation.before);
    //update the end of body range now that the page break has been inserted.
    endOfBodyRange = doc.body.getRange(Word.RangeLocation.end);
    var pageTitleParagraph = endOfBodyRange.insertParagraph('Works Cited', Word.InsertLocation.after);
    pageTitleParagraph.alignment = Word.Alignment.centered;
    var bibRange = pageTitleParagraph.getRange(Word.RangeLocation.after);
    var bibContentControl = bibRange.insertContentControl();
    bibContentControl.title = "Bibliography managed via add-on";
    bibContentControl.tag = "BibliographyControlTag";
    bibContentControl.placeholderText = "Use the add-on to manage your bibliography.";
    bibContentControl.insertText("...loading bibliography...",Word.InsertLocation.end);
    bibContentControl.cannotEdit = true;
    var marketingText = bibRange.insertParagraph("",Word.InsertLocation.after);
    var marketingTextRange = bibRange.getRange(Word.RangeLocation.whole);
    marketingTextRange.insertHtml("<div style='color: #666666; font-size: 10pt;'>"+
        "Bibliography managed by ACME add-on</div>", Word.InsertLocation.after);

    return context.sync().then(function(){
      console.log("Successfully synced");
    }).catch(function(ex){
      console.error('Error: ' + ex.message +' -- ' + JSON.stringify(ex));
      if (ex instanceof OfficeExtension.Error) {
        console.error('  Debug info: ' + JSON.stringify(ex.debugInfo));
      }
    });

  });
}
