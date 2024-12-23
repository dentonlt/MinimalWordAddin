// index.js
//
// minimal Word Addin
//
// small-ish file for use with 'webpack serve' to localhost:3000.
// webpack injects the code below into dist/taskpane.html, then serves
// that file.
// 
// CHANGES
// 20241221 dentonlt
//      respectfully taken/modified the webpack 5 Getting Started:
//      https://webpack.js.org/guides/getting-started/
//      and the MS Office Add-in Sample for Word:
//      https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/hello-world/word-hello-world/taskpane.html
  
/*eslint max-len: ["error", { "code": 80 }]*/

////////// Silly inject example /////////////
function component() {
    const element = document.createElement('div');
    element.innerHTML = "Yo! Cool Add-in, yo!";

    return element;
  }
  document.body.appendChild(component());
    
//// Actual document interaction code //////////////

// appease the error checker:
/* global document, Office, Word */

  Office.onReady((info) => {
      // Check that we loaded into Word
      if (info.host === Office.HostType.Word) {
          document.getElementById("doButton").onclick = callbackDoButton;
      }
  });
  
  function callbackDoButton() {
      return Word.run((context) => {
  
          // insert a paragraph at the start of the document.
          const paragraph = context.document.body.insertParagraph("What a mess ...", Word.InsertLocation.start);
          
          // sync the context to run the previous API call, and return.
          return context.sync();
      });
  }
  