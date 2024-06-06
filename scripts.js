let df = null;
let currentIndex = 0;
let currentUrl = null;

document.getElementById('loadButton').addEventListener('click', () => {
    document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        df = XLSX.utils.sheet_to_json(firstSheet, { raw: true });

        // Añadir columna 'Validacion' con valores vacíos
        df.forEach(record => record['Validacion'] = '');

        // Extraer y actualizar el URL de la fórmula
        df = df.map((record, index) => {
            const cell = firstSheet[`F${index + 2}`]; // Assuming 'Etiqueta' is in column F
            const url = extractUrl(cell ? cell.f : "");
            return { ...record, 'Etiqueta': url };
        });

        document.getElementById('loadButton').disabled = true;
        showHyperlink();
    };
    reader.readAsArrayBuffer(file);
}

function extractUrl(formula) {
    if (formula && typeof formula === 'string') {
        const match = formula.match(/"(https:\/\/plataforma\.rrvsac\.com\/api\/event-media?[^"]+)"/);
        if (match && match[1]) {
            return match[1];
        }
    }
    console.error('No se pudo extraer el URL de la celda:', formula);
    return null;
}

function showHyperlink() {
    if (currentIndex < df.length) {
        const record = df[currentIndex];
        const vehicle = record['Vehículo'];
        const Grupo = record['Grupo primario'];  
        const Alerta = record['Descripción'];
        const Fecha = record['Fecha'];
        const hora = record['Tiempo'];
        const velocidad = record['Velocidad'];
        const url = record['Etiqueta'];
        currentUrl = url;
        document.getElementById('vehicleLabel').innerText = `Placa: ${vehicle}`;
        document.getElementById('Grupo').innerText = `Empresa: ${Grupo}`;
        document.getElementById('Alerta').innerText = `Alerta: ${Alerta}`;
        document.getElementById('Fecha').innerText = `Fecha: ${Fecha}`;
        document.getElementById('hora').innerText = `Hora: ${hora}`;
        document.getElementById('velocidad').innerText = `Velocidad: ${velocidad} km/h`;

        if (isImage(url)) {
            document.getElementById('web-view').style.display = 'none';
            document.getElementById('image-view-container').style.display = 'flex';
            document.getElementById('image-view').style.display = 'block';
            document.getElementById('video-view').style.display = 'none';
            document.getElementById('image-view').src = url;
        } else if (isVideo(url)) {
            document.getElementById('web-view').style.display = 'none';
            document.getElementById('image-view-container').style.display = 'none';
            document.getElementById('video-view').style.display = 'block';
            document.getElementById('video-view').src = url;
        } else {
            document.getElementById('image-view-container').style.display = 'none';
            document.getElementById('video-view').style.display = 'none';
            document.getElementById('web-view').style.display = 'block';
            document.getElementById('web-view').src = url;
        }
        updateCounter();
    } else {
        alert('Selecciones guardadas.');
        document.getElementById('loadButton').disabled = false;
        document.getElementById('downloadButton').style.display = 'block';
        updateCounter();
    }
}

function isImage(url) {
    return (url.match(/\.(jpeg|jpg|gif|png)$/) != null);
}

function isVideo(url) {
    return (url.match(/\.(mp4|webm|ogg)$/) != null);
}

function selectOption(option) {
    df[currentIndex]['Validacion'] = option;
    currentIndex++;
    showHyperlink();
}

function updateCounter() {
    const total = df.length;
    document.getElementById('counterLabel').innerText = `Revisado ${currentIndex} de ${total}`;
}

const optionButtons = document.querySelectorAll('.optionButton');
optionButtons.forEach(button => {
    button.addEventListener('click', () => selectOption(button.innerText));
});

document.getElementById('downloadButton').addEventListener('click', () => {
    const worksheet = XLSX.utils.json_to_sheet(df);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, 'resultados.xlsx');
});
