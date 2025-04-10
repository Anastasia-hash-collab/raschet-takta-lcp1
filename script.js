
let resultText = "";

function calculate() {
  const modules = parseFloat(document.getElementById('numModules').value);
  const palletsPerModule = parseFloat(document.getElementById('palletsPerModule').value);
  const shift = parseFloat(document.getElementById('shiftHours').value);
  const posts = parseFloat(document.getElementById('numPosts').value);
  const tvo = parseFloat(document.getElementById('tvoTime').value);
  const metersPerPallet = parseFloat(document.getElementById('metersPerPallet').value);

  const pallets = modules * palletsPerModule;
  const takt = shift / pallets;
  const taktMin = takt * 60;
  const lineTime = posts * takt;
  const totalCycle = lineTime + tvo;
  const tvoLoadExact = tvo / takt;
  const tvoLoadCeil = Math.ceil(tvoLoadExact);
  const palletsInSystem = posts + tvoLoadCeil;
  const totalMeters = palletsInSystem * metersPerPallet;

  resultText = `
    Интервал подачи/выхода паллет: ${taktMin.toFixed(2)} мин
    Такт системы: ${taktMin.toFixed(2)} мин
    Время прохождения постов: ${lineTime.toFixed(2)} ч
    Загрузка камеры ТВО: ${tvoLoadCeil} паллет (округлено вверх)
    Модулей на выходе в сутки: ${modules.toFixed(2)} модулей
    Требуется борта: ${totalMeters.toFixed(0)} пог. м
    Общий цикл одной паллеты: ${totalCycle.toFixed(2)} ч
  `.trim().replace(/\n/g, "<br>");

  document.getElementById('output').innerHTML = resultText.replace(/\n/g, "<br>");
}

async function downloadDocx() {
  const { Document, Packer, Paragraph } = window.docx;

  const doc = new Document({
    sections: [{
      children: resultText.split("<br>").map(line => new Paragraph(line))
    }]
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "расчет_потока_модулей.docx";
  a.click();
  URL.revokeObjectURL(url);
}
