let express = require('express');
let router = express.Router();
let PizZip = require('pizzip');
let Docxtemplater = require('docxtemplater');
let fs = require('fs');
let path = require('path');

try{
  let ssh = require('node-ssh');
  }catch(e){
    console.log('*** exception cached ***\n' + e);
  }
//Load the docx file as a binary
let content = fs
    .readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');
let zip = new PizZip(content);
let doc = new Docxtemplater();
async function convertDocument(status, name, surname, postal, suburb, state, phone, mobile, email, postcode, fax, x, y) { 
doc.loadZip(zip);
let mr = false;
let mrs = false;
let ms = false;
if(status === "both") both = true;
if(status === "dev") dev = true;
if(status === "test") test = true;
//set the templateVariables
doc.setData({
  both: both,
  dev: dev,
  test: test,
  userid: userid,
  username: username,
  copy: copy,
  email: email
});
try {
  // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
  doc.render()
}
catch (error) {
  // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object).
  let e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
  }
  console.log(JSON.stringify({error: e}));
  if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors.map(function (error) {
          return error.properties.explanation;
      }).join("\n");
      console.log('errorMessages', errorMessages);
      // errorMessages is a humanly readable message looking like this :
      // 'The tag beginning with "foobar" is unopened'
  }
  throw error;
}
let buf = doc.getZip()
           .generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, '..\\output.docx'), buf);
}
/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});
/***
 * Form transformation - inject data into F2 Form happens here
 */
router.post('/form-submit', function(req, res, next) {
  console.log(req.body);
  let status, name, surname, postal, suburb, state, phone, mobile, email, postcode, fax;
  status = req.body.status;
  name = req.body.name;
  surname = req.body.surname;
  postal = req.body.postal;
  suburb = req.body.suburb;
  state = req.body.state;
  phone = req.body.phone;
  mobile = req.body.mobile;
  email = req.body.email;
  postcode = req.body.postcode;
  fax = req.body.fax;
  convertDocument(status, name, surname, postal, suburb, state, phone, mobile, email, postcode, fax);
 // res.render('form-submit', { query: req.body });
});
module.exports = router;
