// index.js
//
// minimal Word Addin
//
// minimal index.js file for use with 'webpack serve' to localhost:3000.
// webpack injects the code below into dist/taskpane.html, then serves
// that file.
// 
// CHANGES
// 20241221 dentonlt - respectfully taken/modified the webpack 5 Getting Started:
//      https://webpack.js.org/guides/getting-started/

function component() {
  const element = document.createElement('div');
  element.innerHTML = "Yo! Cool Add-in, yo!";

  return element;
}
document.body.appendChild(component());
