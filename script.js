document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('upload').addEventListener('change', function (event) {
    cargarExcel(event);
  });
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

    json.forEach((datos, index) => {
      modificarDocumento(datos, index);
    });
  };

  reader.readAsArrayBuffer(file);
}

function modificarDocumento(datos, index) {
  console.log('Datos:', datos);
  const nombreCliente = document.querySelector('p.c12 span');
  if (nombreCliente) nombreCliente.textContent = datos.nombreCliente;

  const nit = document.querySelector('p.c9 span.c4.c0');
  if (nit) nit.textContent = 'RUT No.' + datos.nit;

  const concepto = document.querySelector('p.c7 span.c3');
  if (concepto) concepto.textContent = 'Por concepto de ' + datos.concepto;

  const valor = document.querySelector('p.c7.c15 span.c4.c0');
  if (valor) valor.textContent = `$${datos.valor} (pesos m/cte).`;

  const fechaVencimiento = document.querySelector('p.c2 span.c3');
  if (fechaVencimiento) fechaVencimiento.textContent = `FECHA DE VENCIMIENTO: ${datos.fechaVencimiento}.`;


  setTimeout(() => generarPDF(`documento-modificado-${index}.pdf`), index * 1000);
}

function generarPDF(fileName) {
  html2canvas(document.body).then(canvas => {
    const imgData = canvas.toDataURL('image/png');
    // Accede a jsPDF a través del objeto window.jspdf.jsPDF
    const pdf = new window.jspdf.jsPDF();
    pdf.addImage(imgData, 'PNG', 10, 10);
    pdf.save(fileName);
  }).catch(error => console.error("Error generating PDF", error));
}
