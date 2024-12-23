# minimal word addin project
[uses webpack ... so maybe not so minimal ...]

### CHANGES
20241222 dentonlt - I'm not a programmer. YMMV.

## SUMMARY

This sample project provides a minimal Word Add-in taskpane & server. This
is not "how to create an Add-in," and it is not a viable Add-in
template (this throws an overlay error after a bit). This does, however,
show how Word and the server interact, and how webpack gets custom code to the
taskpane. This also shows an alternative to office-addin-dev-certs for
getting certificates into your project.

This project began with the stock Word Add-in "Getting Started" code and
dropped out all the useful scaffolding (modules for debugging, maintenance,
submission to Azure, etc.).

Provided:
1. a mock manifest.xml. This is used to 'install' the Add-in
  and tell Word the server's URL (in this case, https://localhost:3000)
2. a static html file @ dist/taskpane.html
3. script content from src/index.js webpack'd into dist/main.js
4. https support from localhost:3000 (dummy certs provided)
5. the rest of the webpack/npm settings required.

This was built/tested on Windows 11 with Word Desktop.

## USAGE - GETTING IT RUNNING

You need to:

0. put the project tree somewhere sane.

1. sideload the Add-In to Word. Just "tell Word about manifest.xml".

For Word Desktop:
* copy manifest.xml into a folder where your user has permissions.
  Somewhere like (C:\Users\yourname\Documents\MySecureShareFolder)
* open Word, visit File -> Options -> Trust Center -> Trust Center Settings
  -> Trusted Add-in Catalog
* add your share folder. Provide the path as a URL like
  \\YOURMACHINENAME\path\to\your\sharefolder
* tick the 'show in menu' box for the folder
* close all Office applications, then restart Word
* Open a new document, then ribbon -> Home -> Add-ins -> More Add-ins ->
  Shared Folder -> (this add in)
* if you get a 'cannot read folder' error ... check for typos, or try
  putting the folder somewhere less private (I couldn't under C:\, but
  I could under C:\Users\myname\Documents\)

or, For Word Online: (remove is by clearing browser cache)
* click ribbon -> Home -> Add-ins -> More Add-ins -> My organisation ->
  Upload My Add-in ... then just upload manifest.xml.

2. install npm, node, webpack 5+ & all their trappings.
    Don't forget to get the project modules: `npm install`

3. add the ca.crt file to the Trusted Certificates Store on Windows
* one way: double-click the ca.crt file from Windows File Explorer, add
  to user's Trusted Certificate Store.
* the cert file provided on initial commit will expire in 2025. See below
  for guidelines on making/placing your own certificate files.

4. run the server from command, powershell or WSL: `npm start`
* start webpack, which packs up the node_modules dir and index.js into dist/main.js
* starts 'webpack-dev-server' to serve /dist/template.html

5. in Word, open the new taskpane: ribbon -> Home -> "Pane Button Label"
* check at a browser: https://localhost:3000/template.html
* if the browser works ... Word -should- open the same as a Taskpane. The
   total file is quite large for Word ... so it may take 10-15s or more to be
   available. Unless Word throws an error (certificates, can't find, etc.), it
   may just still be loading. Inspect and/or Refresh.

Since this is running via webpack, changes merged/visible on Refresh. Maybe.


## REMOVING THE ADD-IN

For Word Desktop:
* File -> Options -> Trust Center -> Trusted Add-in Catalogs
* un-tick the folder containing the Add-in manifest.xml
* restart Word
* if this fails ... return to the Trusted Add-in Catalogs and click "Next time
  Office starts, clear all previously-started web add-ins cache." (sorry ...)

For Word Online:
* clear the browser cache.

## OTHER NOTES

package.json
* note the list of required modules, compare to stock

webpack.config.js
* Note Webpack 5+ HTTPS settings. I had no luck with a webpack-generated
  certificate, so this goes the pre-created-certificates route.

src/index.js
* the custom 'stuff' injected into main.js.

Generating Certificates - howto, etc.
1. First, generate ca.crt, ca.key, cert.crt and cert.key
* install mkcert using `npm -i --save-dev mkcert`
* read the mkcert help ... setting certificate properties will help find
   the certificate in the Trusted Certificates Store and bugcheck.
* generate mock certificate authority files:
  `npx mkcert create-ca --organisation "mycaorg" --validity 365`
* generate mock identity files:
  `npx mkcert create-cert --organisation "myserverorg"`
2. Then, combine cert.crt and cert.key into cert.pfx. Easy via PowerShell:
* https://stackoverflow.com/questions/48464705/how-to-get-pfx-file-from-cer-and-key
* `certutil -mergepfx .\source.crt output.pfx`
*  certutil will ask for a password. Pass that to webpack as below.
3. add ca.crt to the Windows Trusted Root Certificates Store
  * From File Explorer, double-click on ca.crt and follow the prompts.
4. provide those certificates to webpack via webpack.config.js:
```
module.exports = {
 ...
  devServer: {
    ...
    },
    server: {
      type: "https",
      options: {
        ca: fs.readFileSync(path.join(__dirname, 'ca.crt')), //CA certificate
        pfx: fs.readFileSync(path.join(__dirname, 'cert.pfx')), // server.pfx 
        key: fs.readFileSync(path.join(__dirname, 'cert.key')),  // server.key
        cert: fs.readFileSync(path.join(__dirname, 'cert.crt')), // server.crt
        passphrase: 'webpack', // password used when creating pfx
        requestCert: false,  
      },
    },
    ...
  },
  ...
};
```



