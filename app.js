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
      documentsStatus.textContent = event.data.message;
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
