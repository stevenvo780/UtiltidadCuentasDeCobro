document.addEventListener('DOMContentLoaded', function () {

  document.getElementById('upload').addEventListener('change', cargarExcel);

  document.getElementById('downloadPDFs').addEventListener('click', generarPDFs);
});

function cargarExcel(event) {
  const input = event.target;
  if (input.files.length === 0) {
    console.log('No file selected.');
    return;
  }
  const file = input.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);

    localStorage.setItem('excelData', JSON.stringify(json));
    document.getElementById('downloadPDFs').disabled = false; // Habilitar el botón para descargar documentos
  };

  reader.readAsArrayBuffer(file);
}

function modificarDocumento(datos, doc) {
  console.log('Datos:', datos);

  const nombreCliente = doc.getElementById('nombre-cliente');
  if (nombreCliente) nombreCliente.textContent = datos.nombreCliente;

  const documentClient = doc.getElementById('documento-cliente');
  if (documentClient) documentClient.textContent = 'RUT No.' + datos.documentClient;

  const concepto = doc.getElementById('concepto');
  if (concepto) concepto.textContent = 'Por concepto de ' + datos.concepto;

  const valor = doc.getElementById('valor-servicio');
  if (valor) valor.textContent = `Valor del servicio $${datos.valor}.`;


  const numeroCuentaCobro = doc.getElementById('numero-cuenta-cobro');
  if (numeroCuentaCobro) numeroCuentaCobro.textContent = `Cuenta de cobro #${datos.numeroCuentaCobro}`;

  const today = new Date();
  const day = today.getDate();
  const month = today.toLocaleDateString('es-ES', { month: 'long' });
  const year = today.getFullYear();

  const fechaActual = doc.getElementById('fecha-actual');
  if (fechaActual) {
    fechaActual.textContent = `Envigado, ${day} de ${month} de ${year}`;
  }

  const mesActual = doc.getElementById('mes-actual');
  if (mesActual) {
    mesActual.textContent = `Mes: ${month}`;
  }

  const fechaVencimientoDate = new Date(datos.fechaVencimiento);
  const dayFechaVencimiento = fechaVencimientoDate.getDate();
  const monthFechaVencimiento = fechaVencimientoDate.toLocaleDateString('es-ES', { month: 'long' });
  const yearFechaVencimiento = fechaVencimientoDate.getFullYear();

  const fechaVencimiento = doc.getElementById('fecha-vencimiento');
  if (fechaVencimiento) fechaVencimiento.textContent = `Fecha de vencimiento: ${dayFechaVencimiento} de ${monthFechaVencimiento} de ${yearFechaVencimiento}`;
}

function generarPDFs() {
  const datos = JSON.parse(localStorage.getItem('excelData'));
  if (datos && datos.length > 0) {
    procesarSiguienteDocumento(0, datos);
  }
}

function procesarSiguienteDocumento(indice, datos) {
  if (indice < datos.length) {
    const data = datos[indice];
    fetch('doc.html')
      .then(response => response.text())
      .then(html => {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        modificarDocumento(data, doc);

        const elementoParaPDF = document.createElement('div');
        document.body.appendChild(elementoParaPDF);

        const styles = doc.querySelectorAll('style');
        let stylesHtml = '';
        styles.forEach(style => {
          stylesHtml += style.outerHTML;
        });

        elementoParaPDF.innerHTML = stylesHtml + doc.body.innerHTML;

        const today = new Date();
        const month = today.toLocaleDateString('es-ES', { month: 'long' });
        generarPDF(`${data.nombreCliente}-${month}.pdf`, elementoParaPDF, function () {
          document.body.removeChild(elementoParaPDF);
          procesarSiguienteDocumento(indice + 1, datos);
        });
      });
  }
}

function generarPDF(fileName, element, callback) {
  element.style.width = '210mm';
  element.style.maxWidth = '210mm';

  html2canvas(element, {
    width: element.offsetWidth,
    windowWidth: element.scrollWidth
  }).then(canvas => {
    const imgData = canvas.toDataURL('image/png');
    const pdf = new window.jspdf.jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    const marginLeft = 10;
    const marginRight = 10;
    const pdfWidth = pdf.internal.pageSize.getWidth() - marginLeft - marginRight;
    const pdfHeight = pdf.internal.pageSize.getHeight();

    const imgProps = pdf.getImageProperties(imgData);
    const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;

    const scaledHeight = imgHeight <= pdfHeight ? imgHeight : pdfHeight;

    pdf.addImage(imgData, 'PNG', marginLeft, 10, pdfWidth, scaledHeight);
    pdf.save(fileName);

    if (typeof callback === "function") {
      callback();
    }
  }).catch(error => console.error("Error generating PDF", error));
}