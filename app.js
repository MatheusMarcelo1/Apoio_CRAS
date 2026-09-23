(function () {
  'use strict';

  const tabs = Array.from(document.querySelectorAll('.tab'));
  const panels = Array.from(document.querySelectorAll('.tool-panel'));
  const folderButton = document.getElementById('select-folder');
  const formatButton = document.getElementById('format-folder');
  const status = document.getElementById('formatter-status');
  const log = document.getElementById('rename-log');
  const supportWarning = document.getElementById('folder-support');
  let selectedDirectory = null;
  const documentsStatus = document.getElementById('documents-status');
  const beneficiaryInput = document.getElementById('beneficiary-csv-input');
  const beneficiarySelect = document.getElementById('select-beneficiary-csv');
  const beneficiarySave = document.getElementById('save-beneficiary-xlsx');
  const beneficiaryStatus = document.getElementById('beneficiary-status');
  const beneficiaryPreview = document.getElementById('beneficiary-preview');
  const beneficiaryTable = document.getElementById('beneficiary-table');
  const beneficiaryShowNis = document.getElementById('beneficiary-show-nis');
  let beneficiaryRows = [];

  function activateTool(tool) {
    tabs.forEach((tab) => {
      const active = tab.dataset.tool === tool;
      tab.classList.toggle('is-active', active);
      tab.setAttribute('aria-selected', String(active));
    });
    panels.forEach((panel) => {
      const active = panel.id === `panel-${tool}`;
      panel.hidden = !active;
    });
  }

  tabs.forEach((tab) => {
    tab.addEventListener('click', () => activateTool(tab.dataset.tool));
  });

  if (beneficiarySelect) {
    beneficiarySelect.addEventListener('click', () => beneficiaryInput.click());
    beneficiaryInput.addEventListener('change', async () => {
      const file = beneficiaryInput.files && beneficiaryInput.files[0];
      if (!file) return;
      try {
        beneficiaryStatus.textContent = 'Lendo e formatando o CSV...';
        beneficiaryRows = formatBeneficiaryCsv(await file.text());
        renderBeneficiaryPreview();
        beneficiarySave.disabled = false;
        beneficiaryStatus.textContent = `${beneficiaryRows.length} beneficiário(s) pronto(s) para exportação.`;
      } catch (error) {
        beneficiarySave.disabled = true;
        beneficiaryStatus.textContent = `Não foi possível formatar o CSV: ${error.message}`;
      }
    });
    beneficiarySave.addEventListener('click', saveBeneficiaryWorkbook);
  }

  if (beneficiaryShowNis) {
    beneficiaryShowNis.addEventListener('change', renderBeneficiaryPreview);
  }

  function parseCsvLine(line) {
    const values = [];
    let value = '';
    let quoted = false;
    for (let index = 0; index < line.length; index += 1) {
      const character = line[index];
      if (character === '"') {
        if (quoted && line[index + 1] === '"') {
          value += '"';
          index += 1;
        } else {
          quoted = !quoted;
        }
      } else if (character === ';' && !quoted) {
        values.push(value);
        value = '';
      } else {
        value += character;
      }
    }
    values.push(value);
    return values;
  }

  function formatBeneficiaryCsv(text) {
    const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter((line) => line.trim());
    if (lines.length < 2) throw new Error('o arquivo não possui linhas de dados.');
    const headers = parseCsvLine(lines[0]).map((header) => header.trim().toUpperCase());
    const required = ['CPF', 'NOME', 'VLRTOTAL', 'SITFAM'];
    const positions = required.map((header) => headers.indexOf(header));
    if (positions.some((position) => position < 0)) {
      throw new Error('as colunas CPF, NOME, VLRTOTAL e SITFAM não foram encontradas.');
    }
    const nisIndex = headers.indexOf('NIS');
    return lines.slice(1).map((line) => {
      const values = parseCsvLine(line);
      const base = positions.map((position) => (values[position] || '').trim());
      const nis = nisIndex >= 0 ? (values[nisIndex] || '').trim() : '';
      return { cpf: base[0], nome: base[1], valor: base[2], situacao: base[3], nis };
    }).filter((row) => parseCurrency(row.valor) !== 0)
      .sort((left, right) => left.nome.localeCompare(right.nome, 'pt-BR', { sensitivity: 'base' }));
  }

  function parseCurrency(value) {
    const normalized = String(value).replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
    const number = Number(normalized);
    return Number.isFinite(number) ? number : 0;
  }

  function renderBeneficiaryPreview() {
    beneficiaryTable.innerHTML = '';
    const head = beneficiaryTable.createTHead().insertRow();
    const showNis = beneficiaryShowNis && beneficiaryShowNis.checked;
    const idLabel = showNis ? 'NIS' : 'CPF';
    [idLabel, 'NOME', 'VLRTOTAL', 'SITFAM'].forEach((label) => {
      const cell = document.createElement('th');
      cell.textContent = label;
      head.appendChild(cell);
    });
    const body = beneficiaryTable.createTBody();
    beneficiaryRows.slice(0, 100).forEach((row) => {
      const tableRow = body.insertRow();
      const idValue = showNis ? row.nis : row.cpf;
      [idValue, row.nome, row.valor, row.situacao].forEach((value) => {
        const cell = tableRow.insertCell();
        cell.textContent = value;
      });
    });
    beneficiaryPreview.hidden = false;
  }

  function saveBeneficiaryWorkbook() {
    if (!beneficiaryRows.length || typeof ExcelJS === 'undefined') {
      beneficiaryStatus.textContent = 'A biblioteca XLSX não foi carregada. Verifique a conexão com a internet.';
      return;
    }
    createBeneficiaryWorkbook().catch((error) => {
      beneficiaryStatus.textContent = `Não foi possível gerar o XLSX: ${error.message}`;
    });
  }

  async function createBeneficiaryWorkbook() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Beneficiários', {
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      views: [{ state: 'frozen', ySplit: 1 }]
    });
    const showNis = beneficiaryShowNis && beneficiaryShowNis.checked;
    const idHeader = showNis ? 'NIS' : 'CPF';
    const idKey = showNis ? 'nis' : 'cpf';
    worksheet.columns = [
      { header: idHeader, key: 'id', width: 18 },
      { header: 'NOME', key: 'nome', width: 42 },
      { header: 'VLRTOTAL', key: 'valor', width: 15, style: { numFmt: '#,##0.00' } },
      { header: 'SITFAM', key: 'situacao', width: 23 }
    ];
    beneficiaryRows.forEach((row) => worksheet.addRow({
      id: row[idKey],
      nome: row.nome,
      valor: parseCurrency(row.valor),
      situacao: row.situacao
    }));
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFB7B7B7' } },
          left: { style: 'thin', color: { argb: 'FFB7B7B7' } },
          bottom: { style: 'thin', color: { argb: 'FFB7B7B7' } },
          right: { style: 'thin', color: { argb: 'FFB7B7B7' } }
        };
      });
    });
    const header = worksheet.getRow(1);
    header.height = 24;
    header.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FF333333' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9E9E9' } };
      cell.alignment = { vertical: 'middle' };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } } };
    });
    worksheet.autoFilter = { from: 'A1', to: `D${beneficiaryRows.length + 1}` };
    worksheet.addConditionalFormatting({
      ref: `A2:D${beneficiaryRows.length + 1}`,
      rules: [{
        type: 'expression',
        formulae: ['$D2="BLOQUEADO"'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' }, fgColor: { argb: 'FFFFC7CE' } },
          font: { color: { argb: 'FF9C0006' } }
        }
      }, {
        type: 'expression',
        formulae: ['$D2="SUSPENSO"'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFEB9C' }, fgColor: { argb: 'FFFFEB9C' } },
          font: { color: { argb: 'FF9C6500' } }
        }
      }]
    });
    worksheet.pageSetup.margins = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
    const output = await workbook.xlsx.writeBuffer();
    const blob = new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    if ('showSaveFilePicker' in window) {
      const handle = await window.showSaveFilePicker({
        suggestedName: 'lista_beneficiarios_formatada.xlsx',
        types: [{ description: 'Planilha Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'lista_beneficiarios_formatada.xlsx';
      link.click();
      URL.revokeObjectURL(link.href);
    }
    beneficiaryStatus.textContent = 'Arquivo XLSX salvo com sucesso.';
  }

  document.querySelectorAll('.tool-panel iframe').forEach((frame) => {
    const resize = () => {
      try {
        const documentBody = frame.contentDocument && frame.contentDocument.body;
        if (documentBody) frame.style.height = `${Math.max(documentBody.scrollHeight + 8, 640)}px`;
      } catch (error) {
        frame.style.height = '1100px';
      }
    };
    frame.addEventListener('load', resize);
    frame.addEventListener('load', () => {
      if (frame.contentDocument && 'ResizeObserver' in window) {
        new ResizeObserver(resize).observe(frame.contentDocument.body);
      }
    });
  });

  window.addEventListener('message', async (event) => {
    if (event.data && event.data.type === 'cras-finder:status') {
      return;
    }
    if (!event.data || event.data.type !== 'cras-finder:pick-directory') return;

    const finderFrame = document.querySelector('#panel-documentos iframe');
    if (!finderFrame || event.source !== finderFrame.contentWindow) return;

    try {
      const directory = await window.showDirectoryPicker({ mode: 'read' });
      event.source.postMessage({
        type: 'cras-finder:directory-selected',
        handle: directory
      }, '*');
    } catch (error) {
      if (error.name !== 'AbortError') {
        event.source.postMessage({
          type: 'cras-finder:directory-error',
          message: error.message
        }, '*');
      }
    }
  });

  const canPickDirectory = 'showDirectoryPicker' in window;
  const canRenameDirectory = canPickDirectory
    && typeof FileSystemDirectoryHandle !== 'undefined'
    && 'move' in FileSystemDirectoryHandle.prototype;
  if (!canPickDirectory || !canRenameDirectory) {
    folderButton.disabled = true;
    supportWarning.style.display = 'block';
  }

  folderButton.addEventListener('click', async () => {
    try {
      selectedDirectory = await window.showDirectoryPicker({ mode: 'readwrite' });
      formatButton.disabled = false;
      status.textContent = `Pasta selecionada: ${selectedDirectory.name}.`;
      log.innerHTML = '';
    } catch (error) {
      if (error.name !== 'AbortError') {
        status.textContent = `Não foi possível selecionar a pasta: ${error.message}`;
      }
    }
  });

  async function getSubdirectories(directory, result, depth) {
    for await (const entry of directory.values()) {
      if (entry.kind !== 'directory') continue;
      result.push({ entry, depth });
      await getSubdirectories(entry, result, depth + 1);
    }
  }

  async function renameFolder(directory) {
    const oldName = directory.name;
    const newName = oldName.toUpperCase();
    if (oldName === newName) return false;
    if (typeof directory.move !== 'function') {
      throw new Error('este navegador não permite renomear pastas pela página. Execute o arquivo Python incluído no projeto.');
    }

    const temporaryName = `.__cras_temp__${crypto.randomUUID()}`;
    await directory.move(temporaryName);
    await directory.move(newName);
    return true;
  }

  formatButton.addEventListener('click', async () => {
    if (!selectedDirectory) return;

    formatButton.disabled = true;
    folderButton.disabled = true;
    log.innerHTML = '';
    status.textContent = 'Lendo subpastas...';

    try {
      const directories = [];
      await getSubdirectories(selectedDirectory, directories, 1);
      directories.sort((a, b) => b.depth - a.depth);
      let renamed = 0;

      for (const item of directories) {
        const nameBefore = item.entry.name;
        if (await renameFolder(item.entry)) {
          renamed += 1;
          const line = document.createElement('li');
          line.textContent = `${nameBefore} → ${nameBefore.toUpperCase()}`;
          log.appendChild(line);
        }
      }

      status.textContent = renamed
        ? `${renamed} pasta(s) renomeada(s) com sucesso.`
        : 'Todas as pastas já estavam em maiúsculas.';
    } catch (error) {
      status.textContent = `A formatação foi interrompida: ${error.message}`;
    } finally {
      folderButton.disabled = false;
      formatButton.disabled = false;
    }
  });
})();
