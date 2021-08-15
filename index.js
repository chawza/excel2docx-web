let functionUrl = "https://btpn-scripts.azurewebsites.net/api/excel2docx-api-trigger";

if (window.location.origin == "file://") {   // local develeopment
    functionUrl = "http://localhost:7071/api/excel2docx-api-trigger";
}

class UploadError extends Error {
    constructor(messsage = "upload error") {
        super(messsage)
        this.name = "UploadError"
    }
}

let fileToDownload = {
    fileLink: '',
    name: ''
}

function getFileToUpload() {
    const fileInput = document.getElementById("file-input");
    const uploadedFile = fileInput.files[0];

    if (uploadedFile == undefined) {
        throw new UploadError("upload file first")
    }
    return uploadedFile;
}

async function UploadFile(file) {
    const URL= `${functionUrl}?filename=${file.name}`
    const respond = await fetch(
        URL,
        {
            method: 'POST',
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Access-Control-Allow-Origin': window.location.origin,
            },
            body: file
        }
    )
    return respond.json();
}

async function handelSubmit() {
    const downloadLabel = document.getElementById('filename-to-download');
    const downloadButton = document.getElementById('download-btn');
    
    try {
        downloadButton.hidden = true;
        downloadLabel.innerHTML = ''
        
        const file = getFileToUpload();

        const responseBody = await UploadFile(file);
        const filename = responseBody['filename'];
        const downloadUrl = `${functionUrl}?filename=${filename}`;

        fileToDownload.fileLink = downloadUrl;
        fileToDownload.name = filename;

        downloadButton.hidden = false;
        downloadLabel.innerHTML = filename
    }
    catch (err) {
        if (err instanceof UploadError) {
            alert(err.message)
        }
        throw err
    }
}

async function handleDownload() {
    const respond = await fetch(
        fileToDownload.fileLink,
        {
            method: 'GET',
            headers: {
                'Access-Control-Allow-Origin': window.location.origin,
            },
        }
    )
    const blobURL = URL.createObjectURL(await respond.blob())

    const anchorDownload = document.createElement('a');
    anchorDownload.display = 'none'
    anchorDownload.href = blobURL;
    anchorDownload.download = fileToDownload.name || 'SS.docx';

    document.body.appendChild(anchorDownload);
    anchorDownload.click()
    document.body.removeChild(anchorDownload)
}

function setup() {
    const downloadButton = document.getElementById('download-btn');
    const submitButton = document.getElementById('submit-button');

    downloadButton.onclick = handleDownload
    document.getElementById('download-btn').hidden = true;

    submitButton.onclick = handelSubmit
}


