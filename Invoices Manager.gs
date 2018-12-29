/**
 * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 * This Script was made to cover the function of creating fodlers for invoices based on the year and week.
 * It runs in JavaScript ES5 - No promisses, no functional programming... 
 *
 * Created by ~ Ricardo Yaben.
 *
 * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 **/


/** ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
*          GLOBALS
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ **/ 

// Array containing the type of files to be downloaded from the attachments (include as many as needed: jpg, png,mk4...)
var fileTypesToExtract = ['pdf', 'csv'];
// Name of the the label where the threads will end up
var labelName_to = 'INVOICES/stored';
// Name of the label that we want to remove after moving the threads
var labelName_from = 'INVOICES/pending';
// Root folder in the GDrive
var folder = 'Ricardo Yaben - Invoices (test)';
// A Date
var date = new Date();

/** ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
*          FUNCTIONS
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ **/

// Returns the week of the year as an integer
function weekOfYear(date){
    var d = new Date(+date);
    d.setHours(0,0,0);
    d.setDate(d.getDate()+4-(d.getDay()||7));
    return Math.ceil((((d-new Date(d.getFullYear(),0,1))/8.64e7)+1)/7);
}

// Returns the final folder where the files will be stored
function getFolder_(folder, date){
    var d = new Date(+date);
    
    var parent_Folder = DriveApp.getRootFolder();
    var child_Folder = exists(parent_Folder, folder);
    var year_Folder = exists(child_Folder, d.getFullYear()+'_'+(d.getMonth() + 1));
    var week_Folder = exists(year_Folder, weekOfYear(d));

    return week_Folder;
}
  
// Checks if the folder already exists, otherwise creates the folder and returns the folder.
function exists(parent, folderName){
    var folder;

    try{
      return folder = parent.getFoldersByName(folderName).next();   
    } catch (folder){
      folder = parent.createFolder(folderName);
      return folder = parent.getFoldersByName(folderName).next();
    }
}

// The muscle - Condesated function that does all the job - Main function!!
function gmailToDrive(){
    //Creates a date using the global one, in case there might be any change to it, so it can be added.
    var d = new Date(+date);      
    //creates the  date formatted like: YYYY/MM/DD // var formattedDate = date.getFullYear()+'/'+(date.getMonth()+1)+'/'+date.getDate(); //Date.getMonth() method goes from 0 to 11!!;
    var formattedDate = d.getFullYear()+'/'+d.getMonth()+'/'+1;

   //build query to search emails
   var query = '';

   //filename:pdf; //'after:'+formattedDate+ //'before:' +formattedDate+ //'newer_than:' +xd (d = days; x = number of days) //... (see documentation on gmail queries)
   for(var i in fileTypesToExtract){
        query += (query === '' ?('filename:'+fileTypesToExtract[i]+' after:'+formattedDate) : (' OR filename:'+fileTypesToExtract[i]+' after:'+formattedDate));
    }
   //fill the query with the inbox and "has: .." if you only want to check the ones with, for example, attachments.
   query = 'label:INVOICES-pending ' + query;

   //collection of threads found
   var threads = GmailApp.search(query);
   var label = getGmailLabel(labelName_to);
   var parentFolder = getFolder_(folder, date);
   var root = DriveApp.getRootFolder();
   for(var i in threads){
     //removes the old label
     threads[i].removeLabel(GmailApp.getUserLabelByName(labelName_from));
       var mesgs = threads[i].getMessages();
           for(var j in mesgs){
               //get attachments
               var attachments = mesgs[j].getAttachments();
               for(var k in attachments){
                   var attachment = attachments[k];
                   var isDefinedType = checkIfDefinedType_(attachment);
                   if(!isDefinedType) continue;
                   var attachmentBlob = attachment.copyBlob();
                   var file = DriveApp.createFile(attachmentBlob);
                   parentFolder.addFile(file);
                   root.removeFile(file);
               }   
           }
       
       threads[i].addLabel(label);
    }  
}

//checks if the attachment extension matches one of the defined
function checkIfDefinedType_(attachment){

    var fileName = attachment.getName();
    var temp = fileName.split('.');
    var fileExtension = temp[temp.length-1].toLowerCase();
    if(fileTypesToExtract.indexOf(fileExtension) !== -1) return true;
    else return false;
}

//gets the new label and if it doesn't exists, it is created
function getGmailLabel(name){
    var label = GmailApp.getUserLabelByName(name);
    if(!label){
        label = GmailApp.createLabel(name);
    }
    return label;
}
