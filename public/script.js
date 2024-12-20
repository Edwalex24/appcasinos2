function uploadFiles() {
    const fileInput1 = document.getElementById('fileInput1');
    const fileInput2 = document.getElementById('fileInput2');
    const file1 = fileInput1.files[0];
    const file2 = fileInput2.files[0];

    // Validación antes de enviar
    if (!file1 || !file2) {
        alert('Por favor selecciona ambos archivos antes de subir.');
        return;
    }

    const formData = new FormData();
    formData.append('file1', file1);
    formData.append('file2', file2);

    // Mostrar indicador de carga
    const uploadButton = document.getElementById('uploadButton');
    uploadButton.disabled = true;
    uploadButton.textContent = 'Subiendo...';

    fetch('/upload', {
        method: 'POST',
        body: formData,
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Error en la respuesta del servidor.');
        }
        return response.blob();
    })
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'Anexo_procesado.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        alert('Archivo descargado con éxito.');
    })
    .catch(error => {
        console.error('Error:', error);
        alert('Hubo un problema al procesar los archivos. Inténtalo de nuevo.');
    })
    .finally(() => {
        // Restaurar estado del botón
        uploadButton.disabled = false;
        uploadButton.textContent = 'Subir Archivos';
    });
}
