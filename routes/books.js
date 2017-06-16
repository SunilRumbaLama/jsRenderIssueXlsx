var express = require('express');
var router = express.Router();
var path = require('path');
var request = require('request');
var fs = require('fs');
var jsreport = require('jsreport-core')({
  tasks: {
    allowedModules: '*'
  }
})

/* GET users listing. */
router.get('/', function (req, res, next) {
  jsreport.init().then(function () {
    jsreport.documentStore.collection('xlsxTemplates')
      .insert({
        // put here the raw content (base64 string) of the xlsx template file, 
        contentRaw: fs.readFileSync(path.resolve('./template.xlsx')).toString('base64'), // this is just an example assuming that the xlsx template is in your directory, you can also use some data that you have sent from other server
        name: 'hello' // name of the new xlsx template file
      }).then(function (doc) {
        console.log(doc.shortid);

        return jsreport.render({
          template: {
            content: fs.readFileSync(path.resolve('./template.xlsx')).toString(),
            xlsxTemplate: {
              shortid: doc.shortid
            },
            engine: 'handlebars',
            recipe: 'xlsx'
          },
          data: {
            "people": [{
              "name": "Alex Rumba",
              "age": 25,
              "gender": "male"
            }, {
              "name": "Roshana Bajracharya",
              "age": 22,
              "gender": "female"
            }]
          },
          options: {
            preview: true
          }
        }).then(function (resp) {
          //prints pdf with headline Hello world 
          resp.stream.pipe(res);
        });

      }).catch(function (e) {
        console.log(e)
      })

  });
});

module.exports = router;