const express = require('express');
const fs = require('fs');
const { promisify } = require('util');
const { nextTick } = require('process');
const writeFile = promisify(fs.writeFile);
const readdir = promisify(fs.readdir);
const url = require('url');
var count='0';


// make sure messages folder exists


 app = express();

app.use( function(req, res, next)  {
  
                app.use(express.static(__dirname + "/public/thankyou.html"));
                res.sendFile(__dirname + "/public/thankyou.html");
});

app.use(express.static('public'));

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Listening on port ${PORT}`);
});
