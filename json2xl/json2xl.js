var xl = rekuire('excel4node');

var Json2XL = function(data, callback){

    function BuildExcel(data) {
      var rows = data.rows || [['Row1_Col1','Row1_Col2',"Row1_Col3"]];
      var config = data.config || null;

       this.WorkSheets = data.worksheets || ['Test'];
       this.FilePath = data.filepath || './exceluploads/';
       this.FileName = data.filename || 'default_template_' + Date.now() + '.xlsx';
       this.Rows = rows;
       this.TotalRows = rows.length;
       this.Config = config;
       this.DataType = "string";
       this.Pattern = "solid";
       this.Color = "white";
       this.BackgroundColor = "none";
       this.FontSize = "10px";
       this.FontFamily = "arial";
       this.FontWeight = "normal";
       this.File = null;
       this.Wb = null; // WorkBook
       this.Ws = null; // Worksheet
       this.WbStyle = null; // WorkBook Style wb.Style();
       this.BorderColor = "black";
    }

    BuildExcel.prototype = {
      getWorkBook: function(){
      	flag = true;
      	var wb;

      	if (this.Config === null) {
      		flag = false;
      	}
      	else if(typeof(this.Config.WbOpts)==='undefined'){
      		flag = false;
      	}

      	wb = (flag===true) ? new xl.WorkBook(this.Config.WbOpts) : new xl.WorkBook();
      	this.Wb = wb;
  		this.WbStyle = wb.Style();
        return wb;
      },
      getWorkSheet: function(){
        	flag = true;
      	var wb = this.Wb;
      	var sheetName = this.WorkSheets[0];

      	if (this.Config === null) {
      		flag = false;
      	}
      	else if(typeof(this.Config.WsOpts)==='undefined'){
      		flag = false;
      	}

      	ws = (flag===true) ? wb.WorkSheet(sheetName, wsOpts) : wb.WorkSheet(sheetName); 
      	this.Ws = ws;
        return wb;
      },
      createDirectory: function(dir){        
         if(!fs.existsSync(dir)){
              fs.mkdirSync(dir);
          }
          return true;
      },
      createCellStyles: function(cell,cellStyle){
        var styleArr = [];
        var borderStyle, borderType, borderColor, borderCoordinates, myStyle;
        var color = (typeof(cellStyle[0].color) === 'undefined') ? this.Color : cellStyle[0].color.toUpperCase();
        var backgroundColor = (typeof(cellStyle[0].backgroundColor) === 'undefined') ? this.BackgroundColor : cellStyle[0].backgroundColor;
        var pattern = (typeof(cellStyle[0].pattern) === 'undefined') ? this.Pattern : cellStyle[0].pattern;
        var fontSize = (typeof(cellStyle[0].fontSize) === 'undefined') ? parseInt(this.FontSize, 10) : parseInt(cellStyle[0].fontSize, 10);
        var fontFamily = (typeof(cellStyle[0].fontFamily) === 'undefined') ? this.FontFamily : cellStyle[0].fontFamily;
        var fontWeight = (typeof(cellStyle[0].fontWeight) === 'undefined') ? this.FontWeight : cellStyle[0].fontWeight;
        var border = (typeof(cellStyle[0].border) === 'undefined') ? null : cellStyle[0].border;
        var freeze = (typeof(cellStyle[0].freeze) === 'undefined') ? "no" : cellStyle[0].freeze.toUpperCase();

        if(color!==null && color!==''){
          styleArr.push(cell.Format.Fill.Color(color));
        }

        if(backgroundColor!==null && backgroundColor!==''){
          styleArr.push(cell.Format.Fill.Pattern(pattern));
          styleArr.push(cell.Format.Fill.Color(backgroundColor));
        }

        if(color!==null && color!==''){
          styleArr.push(cell.Format.Fill.Color(color));
        }

        if(fontSize!==null && fontSize!==''){          
          styleArr.push(cell.Format.Font.Size(8));
        }

        if(fontFamily!==null && fontFamily!==''){
          styleArr.push(cell.Format.Font.Family(fontFamily));
        }

        if(fontWeight!==null && fontWeight!=='' && fontWeight!=='normal'){
          styleArr.push(cell.Format.Font.Bold());
        }

        if(freeze==="YES"){
          styleArr.push(cell.Format.Font.Bold());
        }

        if(border!==null && border.length>0){
              borderType = cellStyle[0].border[0];
              borderColor = (typeof(cellStyle[0].border[1]) === 'undefined') ? this.BorderColor : cellStyle[0].border[1];
              borderCoordinates = (typeof(cellStyle[0].border[2]) === 'undefined') ? null : cellStyle[0].border[2];

              myStyle = this.WbStyle;

              //if(borderCoordinates===null){
                  myStyle.Border({
                    top:{
                        style: borderType,
                        color: borderColor
                    },
                    bottom:{
                        style: borderType,
                        color: borderColor
                    },
                    left:{
                        style: borderType,
                        color: borderColor
                    },
                    right:{
                        style: borderType,
                        color: borderColor
                    }
                  });
              //}
              
              styleArr.push(cell.Style(myStyle));
        }

          return styleArr;
      },
      createRows: function(rows){       
        var row, cell, colsLen, i, j, cellValue, cellType, cellStyle, l;
        var ws = this.Ws;
        var totRows = this.TotalRows;
        var k = 1;
        	flag = true;

        if(this.Config===null || typeof(this.Config.freezePanes)=='undefined'){
        	flag = false;
        }
        
        var freezePanes = (flag===false) ?  null : this.Config.freezePanes;

        for(i=0; i<totRows; i++){            
            row = rows[i];
            colsLen = row.length;
            
            for(j=0; j<colsLen; j++){            	
                cellValue = (typeof row[j].value === 'undefined') ? row[j] : row[j].value;
                cellType = (typeof row[j].dataType === 'undefined') ? 'STRING' : row[j].dataType.toUpperCase();
                cellStyle = (typeof row[j].style === 'undefined') ? null : row[j].style;

                switch(cellType){
                  case 'STRING' : cell = ws.Cell(k, j+1).String(cellValue); break;
                  case 'NUMBER' : cell = ws.Cell(k, j+1).Number(cellValue); break;
                  case 'DATE' : cell = ws.Cell(k, j+1).Date(new Date(cellValue)); break;
                  case 'FORMULA' : cell = ws.Cell(k, j+1).Formula(cellValue); break;
                  case 'LINK' :  
                      if(cellValue instanceof Array) {                  
                        cell = ws.Cell(k, j+1).Link(cellValue[0], cellValue[1]);
                      }
                      else{
                        cell = ws.Cell(k, j+1).Link(cellValue);
                      }
                    break;                  
                }              

                if(cellStyle!=null){
                  this.createCellStyles(cell, cellStyle);
                }
            }
            k++;
        };

        if(freezePanes!==null){
            var freezeRows = (typeof(freezePanes.rows) === 'undefined') ? null : freezePanes.rows;
            var freezeCols = (typeof(freezePanes.cols) === 'undefined') ? null : freezePanes.cols;
            var m, n;

            if(freezeRows!=null){
              for(m=0; m<freezeRows.length; m++){
                ws.Row(freezeRows[m]).Freeze();  
              }              
            }

            if(freezeCols!=null){
              for(n=0; m<freezeCols.length; m++){
                ws.Row(freezeCols[n]).Freeze();  
              }              
            }
        }

        return true;
      },
      createExcelSheet: function(){
          var filePath = this.FilePath;
          var pathExists = this.createDirectory(filePath);

          if(pathExists){
              var wb = this.getWorkBook();
              this.getWorkSheet();

              var rows = this.Rows;
              var rowsCreated = this.createRows(rows);
              var response = {};

              if(rowsCreated){
                var fileName = this.FileName;
                var file = filePath + fileName;
                this.File = file;
                wb.write(file); // Create Excel File                
                response.status = "success";
              }
              else{
              	response.status = "error";                
              }

              response.file = this.FilePath + this.FileName;
              return response;
          }          
      }
    };

    var BuildExcel = new BuildExcel(data);    
    var response = BuildExcel.createExcelSheet();    
    callback(null, response);
}; 

module.exports = {
  Json2XL: Json2XL
}