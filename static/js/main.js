// Upload page — drag-and-drop and multi-file selection
(function () {
    const dropZone     = document.getElementById('dropZone');
    const fileInput    = document.getElementById('fileInput');
    const submitBtn    = document.getElementById('submitBtn');
    const dropContent  = document.getElementById('dropContent');
    const selectedArea = document.getElementById('selectedFilesContent');
    const fileList     = document.getElementById('fileList');
    const form         = document.getElementById('uploadForm');

    if (!dropZone) return; // not on upload page

    function showFiles(files) {
        if (!files || files.length === 0) return;

        const invalid = [...files].filter(f => {
            const ext = f.name.split('.').pop().toLowerCase();
            return !['xls', 'xlsx'].includes(ext);
        });
        if (invalid.length) {
            alert('Formato inválido: ' + invalid.map(f => f.name).join(', ') + '\nUse apenas .xls ou .xlsx');
            return false;
        }

        fileList.innerHTML = '';
        [...files].forEach(f => {
            const chip = document.createElement('span');
            chip.className = 'file-chip';
            chip.textContent = f.name;
            fileList.appendChild(chip);
        });

        dropContent.style.display  = 'none';
        selectedArea.style.display = 'flex';
        dropZone.classList.add('has-file');
        submitBtn.disabled = false;
        return true;
    }

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) showFiles(fileInput.files);
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length) {
            const dt = new DataTransfer();
            [...files].forEach(f => dt.items.add(f));
            fileInput.files = dt.files;
            showFiles(files);
        }
    });

    form.addEventListener('submit', () => {
        document.getElementById('btnText').style.display    = 'none';
        document.getElementById('btnLoading').style.display = 'inline';
        submitBtn.disabled = true;
    });
})();
