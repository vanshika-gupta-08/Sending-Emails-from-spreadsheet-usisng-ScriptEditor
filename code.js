function myFunction() {
  
   
  var originalSS = SpreadsheetApp.openById("1yR5Zi0G0wEeexskn1BopgPcfxAE2TUkf_2aGh6A");
   var sheet1 = originalSS.getSheetByName('Sheet1');
   var filterSS = SpreadsheetApp.openById('1dsT_GIFvty45pAak4dkCR2xxGYgN3C3YjQo0T0E');
   var sheet2 = filterSS.getSheetByName('Sheet1');
    const startRow = 2; // First row of data to process
     const numRows = sheet1.getLastRow();
      const numCols= sheet1.getLastColumn();
   const data = sheet1.getRange("B1:Y1").getValues();
 const wholeNo = data[0][0];
  const  vendorName = data[0][1];
    const dockhitDate = data[0][2];
  const 	DockhitTime = data[0][3];      
   const  gIR  = data[0][4];  	
  const   po	= data[0][5];  
   const  irnNO	= data[0][6];  
   const  invoiceNO	= data[0][7];  
   const  invoice	= data[0][8];  
  const   productDesc	= data[0][9];  
  const   productCode= data[0][10];  
  const  unitPrice	= data[0][11];  
  const   ro	= data[0][12];  
  const   skuID	= data[0][13];  
  const   grnDate 	= data[0][14];  
  const   invoiceQty	= data[0][15];  
 /* const   physicalQty = data[0][16];  
  const   assortmentSize	= data[0][17]; 
   const  accepted = data[0][18]; 
   const  rejected  = data[0][19];  
  const  rejectReason = data[0][20];  
  const   rejectRemark	 = data[0][21];  
  const   excessRemark	= data[0][22];  
  const   grnStatus = data[0][23];
	*/
  
 const htmlelement = HtmlService.createTemplateFromFile("table");
 
  htmlelement.wholeNo =  wholeNo;
   htmlelement.vendorName  = vendorName;
    htmlelement.dockhitDate  = dockhitDate; 
    htmlelement.DockhitTime= DockhitTime;
    htmlelement.gIR= gIR;
      htmlelement.po= po;
    htmlelement.irnNO= irnNO;
    htmlelement.invoiceNO = invoiceNO;
    htmlelement.invoice = invoice;
    htmlelement.productDesc= productDesc;
    htmlelement.productCode= productCode;
    htmlelement.unitPrice= unitPrice;
    htmlelement.ro= ro;
    htmlelement.skuID= skuID;
    htmlelement.grnDate= grnDate;
    htmlelement.invoiceQty= invoiceQty;
    /*  htmlelement.physicalQty= physicalQty;
      htmlelement.assortmentSize= assortmentSize;
      htmlelement.accepted= accepted;
      htmlelement.rejected= rejected;
      htmlelement.rejectReason= rejectReason;
      htmlelement.rejectRemark= rejectRemark;
      htmlelement. excessRemark=  excessRemark;
      htmlelement.grnStatus= grnStatus;
	*/

   getvalues= sheet1.getRange(startRow,1,numRows-1,numCols).getValues();
  var headerValues = sheet1.getRange(1,2,1,numCols-1).getValues();
  sheet2.getRange(1 , 1 , 1 ,numCols-1).setValues(headerValues);  

  for(var i = 0 ; i < numRows-1 ; i++)
{
  var emaiL = getvalues[i][0];
  var name= getvalues[i][2];
  var conD = 1;
  var  l = 2;
 
  for(var j=0; j < i ; j++) 
  {
  if (getvalues[j][0] == emaiL )
     {
       conD = 0;
      break;
     }
  //  break;
    
  }
  if(conD==1)
  {
   
 for(var k= i; k < ( numRows-1)  ;  k++)
 {
   if(getvalues[k][0] == emaiL)
     {
    
       
      var values = sheet1.getRange( k + 2 , 2 , 1 , numCols-1).getValues();  
 
    
      sheet2.getRange(l , 1 , 1 ,numCols-1).setValues(values);  
       
 l=l+1;
     }
 }
  
  }
    if(conD==1)
    {
  
var d= sheet2.getLastRow();
  var tableValues= sheet2.getRange(2,1,d-1,numCols-1).getValues();

for(var time_1= 1; time_1 <= d; time_1++)
     {
  
       var formatedTime =  sheet2.getRange( time_1  , 3).getValue();
       
       var hours = formatedTime.getHours();
  var minutes = formatedTime.getMinutes();
  var ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + '.' + minutes + ' ' + ampm;
        sheet2.getRange( time_1 , 3 ).setValue(strTime);
    }
    
  
  
 htmlelement.tableValues  = tableValues; 
  const htmlEmail = htmlelement.evaluate().getContent();
  var currentDate =  new Date();
      var daTe = currentDate.getDate();
           var month = currentDate.getMonth()+1;
           var year = currentDate.getFullYear();
   var date1 = daTe +"-" + month + "-" + year;
   const Subject ="GRN"+ "/" + name +"/"+ date1 ;
    const  recipientsCC = "vanshikagupta554@gmail.com" ;
   
      
     var spreadsheet =  SpreadsheetApp.openById('1d_GIFvty45pArVqak4dkCR2xxGYgN3C3YjQo0T0E');
       var spreadsheetId = spreadsheet.getId(); 
  var file          = DriveApp.getFileById(spreadsheetId);
  var url           = 'https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
    var fileName = (spreadsheet.getName()) + '.xlsx';
 var blobs   = [response.getBlob().setName(fileName)];
     
      
      MailApp.sendEmail({
  to: emaiL,
  cc: recipientsCC,
  subject: Subject,
  htmlBody: htmlEmail,
 attachments: blobs
});

sheet2.clear();
 sheet2.getRange(1 , 1 , 1 ,numCols-1).setValues(headerValues);   
    }}
 
}