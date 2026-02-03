document.addEventListener('DOMContentLoaded', () => {
    const templateInput = document.getElementById('template');
    const sourceInput = document.getElementById('source');
    const processBtn = document.getElementById('processBtn');
    const extractedDataDiv = document.getElementById('extractedData');
    const dataForm = document.getElementById('dataForm');
    const generateBtn = document.getElementById('generateBtn');
    const statusDiv = document.getElementById('status');

    let templateFile = null;
    let sourceFile = null;
    let extractedData = {};

    templateInput.addEventListener('change', (e) => {
        templateFile = e.target.files[0];
    });

    sourceInput.addEventListener('change', (e) => {
        sourceFile = e.target.files[0];
    });

    processBtn.addEventListener('click', async () => {
        if (!templateFile || !sourceFile) {
            showStatus('Пожалуйста, загрузите оба файла.', 'error');
            return;
        }

        showStatus('Обработка...', 'info');

        try {
            const text = await extractTextFromSource(sourceFile);
            extractedData = parseData(text);
            populateForm(extractedData);
            extractedDataDiv.classList.remove('hidden');
            showStatus('Данные извлечены. Проверьте и скорректируйте при необходимости.', 'success');
        } catch (error) {
            showStatus('Ошибка обработки: ' + error.message, 'error');
        }
    });

    generateBtn.addEventListener('click', async () => {
        const formData = new FormData(dataForm);
        const data = {
            fio: formData.get('fio'),
            property: formData.get('property'),
            address: formData.get('address'),
            cost: formData.get('cost'),
            signature: formData.get('signature')
        };

        showStatus('Генерация договора...', 'info');

        try {
            await generateDocx(templateFile, data);
            showStatus('Договор сгенерирован и скачан.', 'success');
        } catch (error) {
            showStatus('Ошибка генерации: ' + error.message, 'error');
        }
    });

    async function extractTextFromSource(file) {
        if (file.type.startsWith('image/')) {
            // OCR with Tesseract
            const { data: { text } } = await Tesseract.recognize(file, 'rus');
            return text;
        } else if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            // Parse DOCX with Mammoth
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            return result.value;
        } else {
            throw new Error('Неподдерживаемый тип файла источника.');
        }
    }

    function parseData(text) {
        const data = {};

        // ФИО: формат Иванов Иван Иванович
        const fioMatch = text.match(/([А-Я][а-я]+\s[А-Я][а-я]+\s[А-Я][а-я]+)/);
        data.fio = fioMatch ? fioMatch[0] : '';

        // Адрес: искать "г. ", "ул. ", etc.
        const addressMatch = text.match(/(г\.\s*[А-Яа-я]+.*?(ул\.\s*[А-Яа-я]+.*?)?)/);
        data.address = addressMatch ? addressMatch[0] : '';

        // Стоимость: число с запятой
        const costMatch = text.match(/(\d+,\d{2})/);
        data.cost = costMatch ? costMatch[0] : '';

        // Наименование имущества: предположим после адреса до стоимости
        const propertyMatch = text.match(/(?<=\bадрес\b).*?(?=\bстоимость\b|\d+,\d{2})/i);
        data.property = propertyMatch ? propertyMatch[0].trim() : '';

        // Подпись: возможно, не извлекать автоматически
        data.signature = '';

        return data;
    }

    function populateForm(data) {
        document.getElementById('fio').value = data.fio;
        document.getElementById('property').value = data.property;
        document.getElementById('address').value = data.address;
        document.getElementById('cost').value = data.cost;
        document.getElementById('signature').value = data.signature;
    }

    async function generateDocx(templateFile, data) {
        const arrayBuffer = await templateFile.arrayBuffer();
        const zip = new PizZip(arrayBuffer);

        // Извлечь document.xml
        const docXml = zip.file('word/document.xml');
        if (!docXml) throw new Error('Не найден document.xml в шаблоне.');

        let xmlContent = docXml.asText();

        // Парсить XML
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'application/xml');

        // Найти все runs с highlight yellow
        const highlights = [];
        const result = xmlDoc.evaluate('//w:r[w:rPr/w:highlight/@w:val="yellow"]', xmlDoc, nsResolver, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        for (let i = 0; i < result.snapshotLength; i++) {
            highlights.push(result.snapshotItem(i));
        }

        // Функция для разрешения namespaces
        function nsResolver(prefix) {
            const ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            };
            return ns[prefix] || null;
        }

        // Предполагаем порядок: fio, property, address, cost, signature
        const keys = ['fio', 'property', 'address', 'cost', 'signature'];
        highlights.forEach((highlight, index) => {
            if (index < keys.length) {
                const textNode = xmlDoc.evaluate('.//w:t', highlight, nsResolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (textNode) {
                    textNode.textContent = data[keys[index]];
                }
            }
        });

        // Сериализовать обратно
        const serializer = new XMLSerializer();
        xmlContent = serializer.serializeToString(xmlDoc);

        // Обновить ZIP
        zip.file('word/document.xml', xmlContent);

        const out = zip.generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        saveAs(out, 'filled_contract.docx');
    }

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = type;
    }
});