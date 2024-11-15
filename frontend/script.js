// Get DOM elements
const wordToPdfButton = document.getElementById('word-to-pdf');
const pdfToWordButton = document.getElementById('pdf-to-word');
const fileUploadSection = document.getElementById('file-upload-section');
const convertButton = document.getElementById('convert-button');
const resultSection = document.getElementById('result-section');
const resultMessage = document.getElementById('result-message');
const downloadButton = document.getElementById('download-button');
const fileUpload = document.getElementById('file-upload');

// Handle button clicks to show upload section
wordToPdfButton.onclick = () => showUploadSection('word_to_pdf');
pdfToWordButton.onclick = () => showUploadSection('pdf_to_word');

function showUploadSection(convertType) {
    fileUploadSection.style.display = 'block';
    convertButton.style.display = 'inline-block';
    convertButton.onclick = () => convertFile(convertType);
}

function convertFile(convertType) {
    const formData = new FormData();
    const file = fileUpload.files[0];
    if (!file) {
        alert("Please select a file to upload");
        return;
    }

    formData.append('file', file);
    formData.append('convert_type', convertType);

    fetch('http://127.0.0.1:5000/convert', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.error) {
            resultMessage.innerText = data.error;
            resultSection.style.display = 'block';
        } else {
            resultMessage.innerText = data.message;
            resultSection.style.display = 'block';
            const outputFileName = data.file_path.split('/').pop();
            downloadButton.style.display = 'inline-block';
            downloadButton.onclick = () => downloadFile(outputFileName);
        }
    })
    .catch(error => {
        resultMessage.innerText = "An error occurred: " + error;
        resultSection.style.display = 'block';
    });
}

function downloadFile(filename) {
    window.location.href = `http://127.0.0.1:5000/download/${filename}`;
}
