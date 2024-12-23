// webpack.config.js
//
// minimal word addin project
// webpack config for use with 'webpack serve' to localhost:3000
//
// CHANGES
// 20241221 dentonlt - this is really not like the MS Word Add-in sample ...

const path = require('path'); // for safe, cross-platform filename management
const fs = require('fs'); // for fetching certificate files

module.exports = {
  entry: './src/index.js', // list of files to be bundled
  output: {
    filename: 'main.js', // webpack will create this bundle
    path: path.resolve(__dirname, 'dist'), // webpack puts the bundle here
    // don't put "clean: true" here. That will remove the templates from ./dist
    // on build ... which is obviously annoying/bad for this project.
  },
  devServer: {
    port: "3000", // sets server as localhost:3000
    static: { // for serving files (template.html, in particular ...)
      directory: path.join(__dirname, '/dist'), // serve this folder as web root
      watch: true
    },
    server: {
      type: "https", // Word needs an HTTPS connection.
      options: {
        ca: fs.readFileSync(path.join(__dirname, './cert/ca.crt')), //CA certificate
        pfx: fs.readFileSync(path.join(__dirname, './cert/cert.pfx')), // server.pfx 
        key: fs.readFileSync(path.join(__dirname, './cert/cert.key')),  // server.key
        cert: fs.readFileSync(path.join(__dirname, './cert/cert.crt')), // server.crt
        passphrase: 'webpack', // password used when making pfx
        requestCert: false,  // haven't tested true ... hmm ...
      },
    },
  },
};
