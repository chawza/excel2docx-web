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

let loadingSymbol = null;

function showUploadLoading() {
    loadingSymbol = document.createElement('p');
    loadingSymbol.innerHTML = "UPLOADING";
    loadingSymbol.style.color = 'red';

    const submitArea = document.getElementById('submit-area');
    submitArea.appendChild(loadingSymbol);
}

function unshowUploadLoading() {
    document.getElementById('submit-area').removeChild(loadingSymbol);
    loadingSymbol = null;
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

async function UploadAndReturnFile(file) {
    const URL= `${functionUrl}?filename=${file.name}&auto-return=true`;
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
    return await respond.blob();
}

async function handelSubmit() {
    const downloadLabel = document.getElementById('filename-to-download');
    const downloadButton = document.getElementById('download-btn');
    
    try {
        downloadButton.style.display = 'none';
        downloadLabel.innerHTML = '';
        
        showUploadLoading();
        const file = getFileToUpload();

        const responseBody = await UploadFile(file);
        const filename = responseBody['filename'];
        const downloadUrl = `${functionUrl}?filename=${filename}`;

        fileToDownload.fileLink = downloadUrl;
        fileToDownload.name = filename;

        downloadButton.style.display = 'block';
        downloadLabel.innerHTML = filename;
        unshowUploadLoading();
    }
    catch (err) {
        if (err instanceof UploadError) {
            alert(err.message)
        }
        throw err
    }
    finally {
        unshowUploadLoading();
    }
}

function anchorDownloadFile(filename, fileURL) {
    const anchorDownload = document.createElement('a');
    anchorDownload.display = 'none'
    anchorDownload.href = fileURL;
    anchorDownload.download = filename || 'SS.docx';

    document.body.appendChild(anchorDownload);
    anchorDownload.click()
    document.body.removeChild(anchorDownload)
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

    anchorDownloadFile(fileToDownload.name, blobURL);
}

async function handleDownloadSubmit() {
    showUploadLoading(); 
    try {
        const fileToUpload = getFileToUpload();
        const docBlob = await UploadAndReturnFile(fileToUpload);
        const docBlobURL = URL.createObjectURL(docBlob);
    
        let filename = fileToUpload.name;
        if (filename.slice(0,2) == 'TC') {
            filename = 'SS' + filename.slice(2);
        }
        filename = filename.split('.')[0] + '.docx';
        anchorDownloadFile(filename, docBlobURL);
    }
    catch (err) {
        if (err instanceof UploadError) {
            alert(err.message);
        }
        throw err
    }
    finally {
        unshowUploadLoading();
    }
}

function setup() {
    const downloadButton = document.getElementById('download-btn');
    const submitButton = document.getElementById('submit-button');
    const submitDownloadButton = document.getElementById("submit-button-download");

    downloadButton.onclick = handleDownload;
    document.getElementById('download-btn').style.display = 'none';

    submitButton.onclick = handelSubmit;
    submitDownloadButton.onclick = handleDownloadSubmit;
}


