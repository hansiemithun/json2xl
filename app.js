global.rekuire = require("rekuire");
var express = rekuire("express");
var app 	= express();
var json2xl = rekuire('json2xl');
var port    = 8000; 

app.use(express.static(__dirname + '/public'));

app.listen(port, function(err, res){
	if(err){
        throw err;
    }
	else{
        port = process.env.port || port;
        console.log("application running in port " + port);     
    }
});

app.get('/json2xl', function (req, res) {
      var filepath = "uploads/"
      var fileName = Date.now() + '.xlsx';
      
      var wbOpts = {
            jszip:{
                compression:'DEFLATE'
            }
        };

      var wsOpts = {
            margins:{
                left : .75,
                right : .75,
                top : 1.0,
                bottom : 1.0,
                footer : .5,
                header : .5
            },
            printOptions:{
                centerHorizontal : true,
                centerVertical : false
            },
            view:{
                zoom : 100
            },
            outline:{
                summaryBelow : true
            },
            fitToPage:{
                fitToHeight: 100,
                orientation: 'landscape',
          },
        }

      var data = {
              "worksheets" : ['Page-1'],                 
              "filepath": filepath,
              "filename": fileName,                  
              "rows" : [
                             [
                                {   
                                    "value" : "Row-1-Col-1",
                                    "dataType": "string",
                                    "style":[{
                                        "color" : "red",
                                        "backgroundColor" : "green",
                                        "border" : ["thick", "black"]
                                    }]                                       
                                },
                                {   
                                    "value" : "2", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "brown",
                                        "border" : ["thick","blue"]
                                    }]
                                },
                                {   
                                    "value" : "3", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                                {   
                                    "value" : 'B1+C1', 
                                    "dataType": "Formula",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                                {   
                                    "value" : "2015-03-25", 
                                    "dataType": "date",
                                    "style":[{
                                        "color" : "#E6E6E6",
                                        "fontSize" : "58px",
                                        "backgroundColor" : "black"
                                    }]
                                },
                            ],
                            [
                                {   
                                    "value" : "Row-2-Col-1",
                                    "dataType": "string",
                                    "style":[{
                                        "color" : "red"    
                                    }]
                                },
                                {   
                                    "value" : "2", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"                                            
                                    }]
                                },
                                {   
                                    "value" : "3", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                                {   
                                    "value" : 'B2+C2', 
                                    "dataType": "Formula",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                                {   
                                   "value" : ["https://www.google.co.in/", "Google"], 
                                    "dataType": "link",
                                    "style":[{
                                        "color" : "#E6E6E6",
                                        "fontSize" : "58px",
                                        "backgroundColor" : "black"
                                    }]
                                },
                            ],
                       ],                    
              "config" : { 
                    "wbOpts" : wbOpts,
                    "wsOpts" : wsOpts,
                    "freezePanes" : {
                        "rows" : [1],
                        "cols" : [3]
                    }
                }
       };       
});


        app.get('/json2xlmin', function (req, res) {
            var data = {};
            
            json2xl.Json2XL(data, function(err, response){
                    res.json(response);
            });
        });