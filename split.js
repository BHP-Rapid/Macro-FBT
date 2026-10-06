const fs = require('fs');
const html = fs.readFileSync('../prototype.html', 'utf8');

const styleStart = html.indexOf('<style>');
const styleEnd = html.indexOf('</style>');
const css = html.substring(styleStart + 7, styleEnd);
fs.writeFileSync('style.css', css.trim());

const scriptStart = html.lastIndexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const js = html.substring(scriptStart + 8, scriptEnd);
fs.writeFileSync('app.js', js.trim());

let newHtml = html.substring(0, styleStart) + 
  '<link rel="stylesheet" href="style.css">\n' + 
  html.substring(styleEnd + 8, scriptStart) + 
  '<script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>\n<script src="app.js"></script>\n' + 
  html.substring(scriptEnd + 9);
fs.writeFileSync('index.html', newHtml);
