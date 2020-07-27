const express = require('express');
const fs = require('fs');
const { promisify } = require('util');
const { nextTick } = require('process');
const writeFile = promisify(fs.writeFile);
const readdir = promisify(fs.readdir);
const https = require('https');

const options = {
    key: fs.readFileSync('key.pem'),
    cert: fs.readFileSync('cert.pem')
  };
  
  https.createServer(options, function (req, res) {
    res.writeHead(200);
    res.end("hello world\n");
  }).listen(8000);
