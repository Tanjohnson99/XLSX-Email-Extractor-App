document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const downloadButton = document.getElementById('downloadButton');
  const clearButton = document.getElementById('clearButton');
  const emailsContainer = document.getElementById('emailsContainer');
  const emailCount = document.getElementById('emailCount');
  const sidebar = document.getElementById('sidebar');
  const hamburger = document.getElementById('hamburger');
  const sidebarBackdrop = document.getElementById('sidebarBackdrop');
  const navItems = document.querySelectorAll('.nav-item');
  const panels = document.querySelectorAll('.panel');

  const emailSet = new Set();

  downloadButton.addEventListener('click', () => {
    if (emailSet.size === 0) return;
    downloadEmails(emailSet);
  });

  clearButton.addEventListener('click', () => {
    emailSet.clear();
    updateEmailsContainer(emailSet, {
      messageTitle: 'Cleared',
      messageBody: 'Upload new Excel files to extract emails again.'
    });
    setDownloadButtonState(false);
    emailCount.textContent = '0';
    fileInput.value = '';
  });

  fileInput.addEventListener('change', handleFiles);

  navItems.forEach((item) => {
    item.addEventListener('click', () => {
      const targetId = item.dataset.target;
      navItems.forEach((nav) => nav.classList.toggle('active', nav === item));
      panels.forEach((panel) => {
        panel.classList.toggle('active', panel.id === targetId);
      });
      closeSidebar();
    });
  });

  if (hamburger) {
    hamburger.addEventListener('click', () => {
      const isOpen = sidebar.classList.toggle('open');
      hamburger.classList.toggle('active', isOpen);
      toggleBackdrop(isOpen);
      hamburger.setAttribute('aria-expanded', String(isOpen));
    });
  }

  if (sidebarBackdrop) {
    sidebarBackdrop.addEventListener('click', () => {
      closeSidebar();
    });
  }

  function toggleBackdrop(visible) {
    if (!sidebarBackdrop) return;
    if (visible) {
      sidebarBackdrop.hidden = false;
      requestAnimationFrame(() => sidebarBackdrop.classList.add('visible'));
    } else {
      sidebarBackdrop.classList.remove('visible');
      sidebarBackdrop.addEventListener(
        'transitionend',
        () => {
          sidebarBackdrop.hidden = true;
        },
        { once: true }
      );
    }
  }

  function closeSidebar() {
    if (!sidebar.classList.contains('open')) return;
    sidebar.classList.remove('open');
    hamburger.classList.remove('active');
    hamburger.setAttribute('aria-expanded', 'false');
    toggleBackdrop(false);
  }

  async function handleFiles(event) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) {
      return;
    }

    emailSet.clear();
    setDownloadButtonState(false);
    renderProcessingState();

    const supportedFiles = files.filter((file) => file.name.toLowerCase().endsWith('.xlsx'));
    const unsupportedFiles = files.filter((file) => !file.name.toLowerCase().endsWith('.xlsx'));

    if (supportedFiles.length === 0) {
      renderUnsupportedFiles(unsupportedFiles);
      return;
    }

    try {
      await Promise.all(supportedFiles.map(readWorkbook));
      if (emailSet.size > 0) {
        updateEmailsContainer(emailSet);
        emailCount.textContent = String(emailSet.size);
        setDownloadButtonState(true);
      } else {
        updateEmailsContainer(emailSet, {
          messageTitle: 'No emails detected',
          messageBody: 'We scanned the spreadsheets but did not find any email addresses.'
        });
        emailCount.textContent = '0';
      }
      if (unsupportedFiles.length > 0) {
        appendUnsupportedNotice(unsupportedFiles);
      }
    } catch (error) {
      console.error('Failed to read files', error);
      updateEmailsContainer(emailSet, {
        messageTitle: 'Processing error',
        messageBody: 'An error occurred while reading the file(s). Please try again with a valid XLSX file.'
      });
      emailCount.textContent = '0';
    }
  }

  function renderUnsupportedFiles(files) {
    const message = files.length
      ? `Unsupported format for: ${files.map((file) => file.name).join(', ')}. Upload .xlsx spreadsheets to extract emails.`
      : 'Please upload Excel spreadsheets with a .xlsx extension.';
    updateEmailsContainer(emailSet, {
      messageTitle: 'Unsupported files',
      messageBody: message
    });
  }

  function appendUnsupportedNotice(files) {
    if (files.length === 0) {
      return;
    }
    const notice = document.createElement('div');
    notice.className = 'inline-alert';

    const title = document.createElement('strong');
    title.textContent = 'Skipped files:';

    const detail = document.createElement('span');
    const skippedList = files.map((file) => file.name).join(', ');
    detail.textContent = `We ignored unsupported formats for ${skippedList}.`;

    notice.append(title, detail);
    emailsContainer.prepend(notice);
  }

  function renderProcessingState() {
    emailsContainer.replaceChildren(
      createEmptyState('Processing files…', 'Give us a moment while we scan your spreadsheets.')
    );
  }

  function updateEmailsContainer(set, emptyState = null) {
    emailsContainer.innerHTML = '';

    if (emptyState) {
      emailsContainer.appendChild(
        createEmptyState(emptyState.messageTitle, emptyState.messageBody)
      );
      return;
    }

    if (set.size === 0) {
      emailsContainer.appendChild(
        createEmptyState('No emails yet', 'Upload one or more Excel files to see the extracted addresses.')
      );
      return;
    }

    const emailsList = document.createElement('ul');
    const sortedEmails = Array.from(set).sort((a, b) => a.localeCompare(b));

    sortedEmails.forEach((email) => {
      const listItem = document.createElement('li');
      listItem.textContent = email;
      emailsList.appendChild(listItem);
    });

    emailsContainer.appendChild(emailsList);
  }

  function createEmptyState(title, body) {
    const wrapper = document.createElement('div');
    wrapper.className = 'empty-state';

    const heading = document.createElement('h4');
    heading.textContent = title;

    const paragraph = document.createElement('p');
    paragraph.textContent = body;

    wrapper.append(heading, paragraph);
    return wrapper;
  }

  function setDownloadButtonState(isEnabled) {
    downloadButton.disabled = !isEnabled;
  }

  function readWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (loadEvent) => {
        try {
          const arrayBuffer = loadEvent.target.result;
          const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
          workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              raw: false,
              blankrows: false
            });
            extractEmailsFromSheetData(sheetData);
          });
          resolve();
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('Unable to read file'));
      reader.readAsArrayBuffer(file);
    });
  }

  function extractEmailsFromSheetData(sheetData) {
    const emailRegex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/gu;

    sheetData.forEach((row) => {
      row.forEach((cell) => {
        if (typeof cell !== 'string') {
          return;
        }
        const normalized = cell.trim();
        if (!normalized) {
          return;
        }
        const matches = normalized.match(emailRegex);
        if (matches) {
          matches.forEach((email) => emailSet.add(email.toLowerCase()));
        }
      });
    });
  }

  function downloadEmails(set) {
    const emailsText = Array.from(set).sort((a, b) => a.localeCompare(b)).join('\n');
    const blob = new Blob([emailsText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'extracted_emails.txt';
    anchor.click();
    URL.revokeObjectURL(url);
  }
});
