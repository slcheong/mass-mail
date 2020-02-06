// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// No Node.js APIs are available in this process because
// `nodeIntegration` is turned off. Use `preload.js` to
// selectively enable features needed in the rendering
// process.
const { ipcRenderer, remote } = require('electron');
const Excel = require('exceljs');

localStorage.clear();

// editor add function to select variable
Jodit.defaultOptions.controls.customVariable = {
    iconURL: "../src/edit.png",
    popup: function (editor) {
        let selection = '<select id = "customVar" style="padding: 10px;"> <option>Select Variable</option>' 
        let varName = JSON.parse(localStorage.getItem("varName")) || [];
        for (i = 0; i <varName.length; i++) {
            selection += `<option value = ${varName[i]}> ${varName[i]} </option>`
        }
        selection += "</select>"

        let display = editor.create.fromHTML(selection);
        display.addEventListener("change",() => {
            editor.selection.insertHTML(document.getElementById('customVar').value);
        })

        let message = '<div style = "width: 100px">Upload a valid excel file before select</div>'

        if (varName.length == 0) {
            return message;
        }
        else {
            return (display);
        }
    },
};

// editor add function to embed image
Jodit.defaultOptions.controls.customImage = {
    iconURL: "../src/image.png",
    popup: function (editor) {

        let imagePath = localStorage.getItem("imagePath") || ""
        let selection = `<image src = "${imagePath}" />`
        let message = '<div style = "width: 100px">Select an image to upload before embeded</div>'

        if (imagePath == "") {
            return message;
        }
        else {
            editor.selection.insertHTML(selection);
        }
    },
};

// make text area a jodit editor
var editor = new Jodit('#editor',{
    extraButtons: ["customVariable", "customImage"]
});


// function to check excel
validateExcel = () => {
    try {
        let fileName = document.getElementById("excelFile").files[0].name;
        try {
            fileType = fileName.split(".").pop();
            if (fileType != "xlsx")
                {
                    document.getElementById("excelFile").value = null
                    return "Only Excel File (xlsx) is allowed.";
                }
        } catch (error) {
            return  "There is an error."
        }
    } catch (e) {
        return  "Please upload an excel file";
    }
    return "OK"
}

// function to check form
validateForm = () => {
    let result = validateExcel();
    if (result == "OK")
    {
        try {
            if(document.getElementById("title").value == ""){
                return  "Email Title cannot be empty"
            }
            else if (document.getElementById("editor").value == ""){
                return  "Email content cannot be empty"
            }
        } catch (error) {
            return  "There is an error."
        }
        return "OK"
    }
    else{
        return result
    }

}

// function to check image type
validateImage = (element) => {
    if(typeof element.files[0] != "undefined") {
        imageName = element.files[0].name;
        imagePath = element.files[0].path;
        let allowedImageType = ["jpeg","png","jpg"]
        try {
            imageType = imageName.split(".").pop();
            console.log(imageType)
            if (allowedImageType.includes(imageType.toLowerCase()))
                {
                    return "OK";
                }
            else{
                element.value = null
                return "Pease only append png or jpeg image"
            }
        } catch (error) {
            return  "There is an error."
        }
       }
    else {
        return "No Image"
    }
}


extractExcel = (path) => {
    var workbook = new Excel.Workbook();
    var content = []
    var toEmail = []
    var varName = []
    workbook.xlsx.readFile(path).then(function (){
        var worksheet = workbook.getWorksheet(1)

        // push content
         for( j =1; j < worksheet.rowCount; j++){
             var htmlText = document.getElementById("editor").value;
             for (var i = "B".charCodeAt(); i < "A".charCodeAt()+worksheet.columnCount;  i++)
             {
                 htmlText = htmlText.split(`{${worksheet.getCell(`${String.fromCharCode(i)}1`).value}}`).join(worksheet.getCell(`${String.fromCharCode(i)}${j+1}`).value);
             }
             content.push(htmlText);
        }

        // push email add
        for( j=1; j< worksheet.rowCount; j++){
            if (worksheet.getCell(`A${j+1}`).value != null) {
                toEmail.push(worksheet.getCell(`A${j+1}`).value.text);
            }        
        }
         
        // push variable name
         for (var i = "B".charCodeAt(); i < "A".charCodeAt()+worksheet.columnCount;  i++)
         {
             varName.push(`{${worksheet.getCell(`${String.fromCharCode(i)}1`).value}}`);
         }

         //store info to local storage
         localStorage.setItem("content",JSON.stringify(content));
         localStorage.setItem("email",JSON.stringify(toEmail));
         localStorage.setItem("varName",JSON.stringify(varName));
     })

}

returnError = (msg = null) =>
    {
        let errorMessage = document.getElementById("error");
        if (msg == null )
            {
                errorMessage.classList.remove("alert");
                errorMessage.classList.remove("alert-danger");
                errorMessage.innerHTML = ""
            }
        else
            {
                errorMessage.classList.add("alert");
                errorMessage.classList.add("alert-danger");
                errorMessage.innerHTML = `${msg}`
            }
    }

document.querySelector('#excelFile').addEventListener('change', function(e){
    e.preventDefault();
    returnError();
    let result = validateExcel()
    if(result == "OK")
    {
        try {
            let excelFile = document.getElementById("excelFile").files[0]
            let path = (typeof excelFile == "undefined") ? "" : excelFile.path
            extractExcel(path);
        }
        catch (err) {
            returnError(err)
        }

    }
    else{
        returnError(result);
    }


});

document.querySelector('#image').addEventListener('change', function(e){
    e.preventDefault();
    let imageElement = document.getElementById("image");
    let result = validateImage(imageElement)
    if(result == "OK")  {
        localStorage.setItem("imageName",imageElement.files[0].name);
        localStorage.setItem("imagePath",imageElement.files[0].path);
        returnError();
    }
    else if (result == "No Image"){
        localStorage.removeItem("imageName");
        localStorage.removeItem("imagePath");
    }
    else{
        returnError(result);
    }
});

document.querySelector('#form1').addEventListener('submit', function(e){
    e.preventDefault();
    let imageElement = document.getElementById("image");
    let excelFile = document.getElementById("excelFile").files[0]
    let path = (typeof excelFile == "undefined") ? "" : excelFile.path
    let result = validateForm();
    if (result == "OK"){
        returnError();
        localStorage.setItem("title",document.getElementById("title").value);
        extractExcel(path);
        let imageError = validateImage(imageElement);
        console.log(imageError)
        if (imageError == "OK")
        {
            
            localStorage.setItem("imageName",imageElement.files[0].name);
            localStorage.setItem("imagePath",imageElement.files[0].path);
            
        }
        else if (imageError == "No Image")
        {
            localStorage.removeItem("imageName");
            localStorage.removeItem("imagePath");
        }
        else{
            returnError(imageError);
            return
        }
    
       if(typeof document.getElementById("attachment").files[0] != "undefined") {
        attachmentName = document.getElementById("attachment").files[0].name;
        attachmentPath = document.getElementById("attachment").files[0].path;
        localStorage.setItem("attachmentName",attachmentName);
        localStorage.setItem("attachmentPath",attachmentPath);
       };
        ipcRenderer.send("hideWindow",false);
    }
    else 
    {
        returnError(result);
    }
})

document.getElementById("title").addEventListener("keydown",() => {
    returnError();
})

document.getElementById("editor").addEventListener("change",() => {
    returnError();
})



