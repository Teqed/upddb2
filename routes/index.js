var express = require('express');
var router = express.Router();

var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

// Load the docx file as a binary
var content = fs
    .readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');

var zip = new PizZip(content);

var doc = new Docxtemplater();

async function convertDocument(status, name, surname, postal, suburb, state, phone, mobile, email, postcode, fax, x, y) {

doc.loadZip(zip);

var mr = false;
var mrs = false;
var ms = false;

if(status === "mr") mr = true;

if(status === "mrs") mrs = true;

if(status === "ms") ms = true;

// set the templateVariables
doc.setData({
  hasMr: mr,
  hasMrs: mrs,
  hasMs: ms,
  name: name,
  surname: surname,
  postal: postal,
  suburb: suburb,
  state: state,
  phone: phone,
  mobile: mobile,
  email: email,
  postcode: postcode,
  fax: fax
});

try {
  // renders the document (replacing all occurences of {first_name} by James, {last_name} by Smith, ...)
  doc.render()
}
catch (error) {
  var e = {
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
  }
  throw error;
}

var buf = doc.getZip()
           .generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, '..\\output.docx'), buf);

}

/* GET the home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

/***
 * Form transformation - inject data into a word document
 */
router.post('/form-submit', function(req, res, next) {

  var status, name, surname, postal, suburb, state, phone, mobile, email, postcode, fax;

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

});

module.exports = router;
