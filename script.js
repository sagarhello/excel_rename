document.addEventListener('DOMContentLoaded', () => {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const fileList = document.getElementById('file-list');
  const excludedList = document.getElementById('excluded-list');

  let renamedFiles = []; // Move to global scope
  let excludedFiles = []; // Move to global scope

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropZone.style.backgroundColor = '#f0f8ff';
  });
  dropZone.addEventListener('dragleave', () => {
    dropZone.style.backgroundColor = 'white';
  });
  dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropZone.style.backgroundColor = 'white';
    processFiles(event.dataTransfer.files);
  });
  fileInput.addEventListener('change', (event) => {
    processFiles(event.target.files);
  });

  function processFiles(files) {
    fileList.innerHTML = '';
    excludedList.innerHTML = '';

    renamedFiles = []; // Reset the array
    excludedFiles = []; // Reset the array

    let filesProcessed = 0; // Track processed files

    Array.from(files).forEach((file) => {
      if (!file.name.endsWith('.xlsx')) {
        excludedFiles.push({ name: file.name, reason: 'Invalid file type' });
        updateExcludedList();
        checkAllProcessed(files.length, ++filesProcessed);
        return;
      }

      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let sapCode = 'UNKNOWN';
        let branchCode = 'UNKNOWN';

        jsonData.forEach((row) => {
          let rowData = row.join(' ');
          if (rowData.includes('SAP Code')) {
            sapCode = row[1] || 'UNKNOWN';
          }
          if (rowData.includes('BRANCH CODE')) {
            branchCode = row[1]?.split(' ')[0] || 'UNKNOWN';
          }
        });

        if (sapCode === 'UNKNOWN' || branchCode === 'UNKNOWN') {
          excludedFiles.push({
            name: file.name,
            reason: 'Missing SAP Code or Branch Code',
          });
        } else {
          let newFileName = `${branchCode}_${sapCode}_${file.name}`;
          renamedFiles.push({ file, newFileName });
        }

        updateExcludedList();
        updateRenamedList();
        checkAllProcessed(files.length, ++filesProcessed);
      };
    });
  }

  function checkAllProcessed(totalFiles, processedCount) {
    if (processedCount === totalFiles) {
      console.log(`All ${totalFiles} files processed.`);
      if (renamedFiles.length > 0) {
        startBatchDownload(); // Only start download after all files are processed
      }
    }
  }

  function updateRenamedList() {
    fileList.innerHTML = '';
    renamedFiles.forEach(({ newFileName }) => {
      let li = document.createElement('li');
      li.className = 'success';
      li.innerText = newFileName;
      fileList.appendChild(li);
    });
  }

  function updateExcludedList() {
    excludedList.innerHTML = '';
    excludedFiles.forEach(({ name, reason }) => {
      let li = document.createElement('li');
      li.className = 'error';
      li.innerText = `${name} - ${reason}`;
      excludedList.appendChild(li);
    });
  }

  function startBatchDownload() {
    let index = 0;
    function downloadNext() {
      if (index < renamedFiles.length) {
        let { file, newFileName } = renamedFiles[index];
        downloadFile(file, newFileName);
        index++;
        setTimeout(downloadNext, 300); // 300ms delay to prevent browser blocking
      }
    }
    downloadNext();
  }

  function downloadFile(file, newFileName) {
    const url = URL.createObjectURL(file);
    const link = document.createElement('a');
    link.href = url;
    link.download = newFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
});
