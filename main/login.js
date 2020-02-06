const { ipcRenderer, remote } = require('electron');
const nodemailer = require("nodemailer");


var email = JSON.parse(localStorage.getItem("email"));
var content = JSON.parse(localStorage.getItem("content"));
var title = localStorage.getItem("title");

emailPreview = (num) => {
    let previewArea = document.getElementById("sample");
    let previewMessage = `<p> To: ${email[num]}</p> <p> Title: ${title}</p> <div class="container">${content[num]}</div>`
    previewArea.innerHTML = previewMessage
}



document.getElementById("sampleSelector").setAttribute("max", email.length);
document.getElementById("sampleSelector").setAttribute("value", 1);
document.getElementById("totalEmail").innerHTML = email.length;
emailPreview(1);


document.querySelector('#close').addEventListener("click", function(){
    ipcRenderer.send("hideWindow",true)
})

document.querySelector('#form2').addEventListener("submit", function(e){
    e.preventDefault();

    let emailAdd = document.getElementById("emailAdd").value;
    let password = document.getElementById("pswd").value;
    let test = validateField(emailAdd, password);
    if (test == "OK")
        {
            sendemail(emailAdd,password);
        }
    else
        {
            returnError(test)
        }
    
});

document.querySelector('#form2').addEventListener("keydown", function(e){
    returnError();

});


document.querySelector('#sampleSelector').addEventListener("change", function(e){
    e.preventDefault();
    sample = document.getElementById("sampleSelector").value - 1;
    emailPreview(sample)
});



validateField = (email,pass) =>{
    if (email == "")
        {
            return "Email cannot be empty"
        }
    else if ( pass == "")
        {
            return "Password cannot be empty"
        }
    else    
        {
            return "OK"
        }
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


toggleButton = (hide) =>{
    let submitButton = document.getElementById("submit")
    let closeButton = document.getElementById("close")
    if( hide == true) {
        submitButton.setAttribute("disabled","")
        submitButton.innerHTML = "Sending Email..."
        closeButton.setAttribute("disabled","")
    }
    else{
        submitButton.removeAttribute("disabled")
        submitButton.innerHTML = "Send Email"
        closeButton.removeAttribute("disabled")
    }
}

async function sendemail(emailAdd, password) {
  
    toggleButton(true)
    let sentMail = []

    const transporter = nodemailer.createTransport({
        pool: true,
        host: "smtp.gmail.com",
        port: 587,
        secure: false, // use TLS
        auth: {
          user: emailAdd,
          pass: password
        }
      });
    
    let verifyLogin = new Promise ( (resolve, reject) => {
        transporter.verify(function(error, success) {
            if (error) {
              returnError(error)
              toggleButton(false)
              return "Error " + error
            } else {
                console.log("Server is ready to take our messages");
                resolve("Server is ready to take our messages");
            }
          });
    
    })

    verifyLogin.then( () => {
        console.log("then start");
        let imageCID = "image1@workingelf";
        let imagePath = localStorage.getItem("imagePath") || "";
        let imageName = localStorage.getItem("imageName");
        let attachmentName = localStorage.getItem("attachmentName")
        let attachmentPath = localStorage.getItem("attachmentPath") || ""
        let email = JSON.parse(localStorage.getItem("email"));
        let content = JSON.parse(localStorage.getItem("content"));
        let title = localStorage.getItem("title");
    
        for (i = 0; i < email.length; i++){
            let last = false;
            if (i+1 == email.length) {last = true;}

            var mailOptions = {
                from: emailAdd, // sender address
                to: email[i], // list of receivers
                subject: title, // Subject line
                html: content[i],
            };
    
            if (imagePath != ""){
                mailOptions.attachments = []
                mailOptions.attachments.push({filename:imageName, path:imagePath, cid: imageCID})
                mailOptions.html = mailOptions.html.replace(imagePath,`cid:${imageCID}`)
            }

            if (attachmentPath != ""){
                if (imagePath == ""){
                    mailOptions.attachments = []
                }
                mailOptions.attachments.push({filename:attachmentName, path:attachmentPath})
            }
    
            if ( i == 0 ) {
                let log = document.getElementById("result")
                let p = document.createElement("p")
                let text = document.createTextNode('Sending Emails...')
                log.classList.add("alert");
                log.classList.add("alert-primary");
                p.appendChild(text)
                log.appendChild(p);
    
            };

            transporter.sendMail(mailOptions, (error, info) => {
                let log = document.getElementById("result")
                let p = document.createElement("p")
                if(error) {
                    let text = document.createTextNode(`Failed to sent to ${info.rejected}. ${error}.`)
                    p.appendChild(text)
                    log.appendChild(p);
                    console.log(info)
                }
                else {
                    console.log(info)
                    let text = document.createTextNode(`Sent to ${info.accepted} successful.`);
                    p.appendChild(text)
                    log.appendChild(p);
                    console.log(info)
                }
                if(last)
                {
                    toggleButton(false)
                }
            });
        }
    })
}
