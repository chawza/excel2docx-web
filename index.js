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

async function UploadAndReturnFile(file, uac_tc = false) {
    const URL= `${functionUrl}?filename=${file.name}&auto-return=true&uac_tc=${uac_tc}}`;
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
        const isChecked = document.getElementById('uac-tc').checked;
        const fileToUpload = getFileToUpload();
        const docBlob = await UploadAndReturnFile(fileToUpload, isChecked);
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
    const submitDownloadButton = document.getElementById("submit-button-download");
    const uploadMessage = document.getElementById('upload-message');
    const fileInput = document.getElementById('file-input');

    downloadButton.onclick = handleDownload;
    document.getElementById('download-btn').style.display = 'none';

    submitDownloadButton.onclick = handleDownloadSubmit;


    fileInput.onchange = function () {
        uploadMessage.innerHTML = fileInput.value.split("\\").pop();
    };
}


