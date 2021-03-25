const fs = require("fs")
const Zip = require('jszip');
const https = require("https");
const Stream = require("stream").Transform;
 
//const ImageModule = require("./es6");
var DocxTemplater = require('docxtemplater');

//Node.js example
var ImageModule = require('docxtemplater-image-module-free');

var content = fs.readFileSync('./template.docx', 'binary');

const data = {
    image1: "https://docxtemplater.com/xt-pro.png",
    image2: "https://images.freeimages.com/images/large-previews/b3d/flowers-1375316.jpg"
  };

//Below the options that will be passed to ImageModule instance
var opts = {}
opts.centered = false; //Set to true to always center images
opts.fileType = "docx"; //Or pptx


//Pass your image loader
opts.getImage = function(tagValue, tagName) {
    console.log(tagValue, tagName);
    // tagValue is "https://docxtemplater.com/xt-pro-white.png" and tagName is "image"
    return new Promise(function (resolve, reject) {
      getHttpData(tagValue, function (err, data) {
        if (err) {
          return reject(err);
        }
        resolve(data);
      });
    });
}
 
//Pass the function that return image size
opts.getSize = function(img, tagValue, tagName) {
    console.log(tagValue, tagName);
    // img is the value that was returned by getImage
    // This is to force the width to 600px, but keep the same aspect ration
    const sizeOf = require("image-size");
    const sizeObj = sizeOf(img);
    console.log(sizeObj);
    const forceWidth = 300;
    const ratio = forceWidth / sizeObj.width;
    return [
      forceWidth,
      // calculate height taking into account aspect ratio
      Math.round(sizeObj.height * ratio),
    ];
}
 
const imageModule = new ImageModule(opts);
 
const zip = new Zip(content);
const doc = new DocxTemplater()
  .loadZip(zip)
  .attachModule(imageModule)
  .compile();
 
doc
  .resolveData(data)
  .then(function () {
    console.log("data resolved");
    doc.render();
    const buffer = doc
      .getZip()
      .generate({
        type: "nodebuffer",
        compression: "DEFLATE"
      });
 
    fs.writeFileSync("./test.docx", buffer);
    console.log("rendered");
  })
  .catch(function (error) {
    error.properties.errors.forEach(function(err) {
        console.log(err);
    });
  });
 
  
function getHttpData(url, callback) {
  https
    .request(url, function (response) {
      if (response.statusCode !== 200) {
        return callback(
          new Error(
            `Request to ${url} failed, status code: ${response.statusCode}`
          )
        );
      }
 
      const data = new Stream();
      response.on("data", function (chunk) {
        data.push(chunk);
      });
      response.on("end", function () {
        callback(null, data.read());
      });
      response.on("error", function (e) {
        callback(e);
      });
    })
    .end();
}






/*
var imageModule = new ImageModule(opts);
var zip = new Zip(content);
var doc = new Docxtemplater()
    .attachModule(imageModule)
    .loadZip(zip)
    .setData({image: './bon point.jpg'})
    .render();
 
var buffer = doc
        .getZip()
        .generate({type:"nodebuffer"});
 
fs.writeFileSync("./test.docx",buffer);*/