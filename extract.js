document.addEventListener('DOMContentLoaded', () => {
    const documentInput = document.getElementById('document');
    const extractBtn = document.getElementById('extractBtn');
    const resultDiv = document.getElementById('result');
    const textOutput = document.getElementById('textOutput');
    const statusDiv = document.getElementById('status');

    let docFile = null;

    documentInput.addEventListener('change', (e) => {
        docFile = e.target.files[0];
    });

    extractBtn.addEventListener('click', async () => {
        if (!docFile) {
            showStatus('Пожалуйста, загрузите документ.', 'error');
            return;
        }

        showStatus('Извлечение текста...', 'info');

        try {
            const text = await extractAllText(docFile);
            textOutput.value = text;
            resultDiv.classList.remove('hidden');
            showStatus('Текст извлечен.', 'success');
        } catch (error) {
            showStatus('Ошибка: ' + error.message, 'error');
        }
    });

    async function extractAllText(file) {
        if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            const arrayBuffer = await file.arrayBuffer();
            const zip = new PizZip(arrayBuffer);

            // Извлечь обычный текст
            const result = await mammoth.extractRawText({ arrayBuffer });
            let text = result.value;

            // Извлечь текст из изображений
            const imageTexts = await extractImagesAndOCR(zip);
            if (imageTexts.length > 0) {
                text += '\n\n--- Текст из изображений ---\n' + imageTexts.join('\n---\n');
            }

            return text;
        } else {
            throw new Error('Поддерживается только DOCX.');
        }
    }

    async function extractImagesAndOCR(zip) {
        const imagePromises = [];
        const mediaFolder = zip.folder('word/media');
        if (mediaFolder) {
            mediaFolder.forEach((relativePath, file) => {
                if (relativePath.match(/\.(png|jpg|jpeg|gif|bmp)$/i)) {
                    const imageBlob = new Blob([file.asArrayBuffer()], { type: 'image/' + relativePath.split('.').pop() });
                    imagePromises.push(Tesseract.recognize(imageBlob, 'rus').then(({ data: { text } }) => text.trim()));
                }
            });
        }
        return await Promise.all(imagePromises);
    }

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = type;
    }
});