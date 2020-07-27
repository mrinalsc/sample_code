const express = require('express');
const fs = require('fs');
const { promisify } = require('util');
const { nextTick } = require('process');
const writeFile = promisify(fs.writeFile);
const readdir = promisify(fs.readdir);
var http=require('http')
var server=http.createServer((function(request,response)
{
	response.writeHead(200,
	{"Content-Type" : "text/plain"});
	response.end("thankyou for checking the code");
}));

const PORT = process.env.PORT || 3001;
server.listen(PORT, () => {
  console.log(`Listening on port ${PORT}`);
});
