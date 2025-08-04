
// Enhanced floating particles with better animation
function createParticles() {
    const particlesContainer = document.getElementById('particles');
    const particleCount = Math.floor(window.innerWidth / 20); // Responsive count
    
    for (let i = 0; i < particleCount; i++) {
      const particle = document.createElement('div');
      particle.classList.add('particle');
      
      // Random size between 1px and 3px
      const size = Math.random() * 2 + 1;
      particle.style.width = `${size}px`;
      particle.style.height = `${size}px`;
      
      // Random position
      particle.style.left = `${Math.random() * 100}%`;
      particle.style.top = `${Math.random() * 100}%`;
      
      // Random animation duration and delay
      const duration = Math.random() * 20 + 10;
      const delay = Math.random() * 20;
      particle.style.animationDuration = `${duration}s`;
      particle.style.animationDelay = `${delay}s`;
      
      // Random opacity
      particle.style.opacity = Math.random() * 0.4 + 0.2;
      
      // Random blur
      particle.style.filter = `blur(${Math.random() * 1}px)`;
      
      particlesContainer.appendChild(particle);
    }
  }
  
  // Initialize particles when DOM is loaded
  document.addEventListener('DOMContentLoaded', function() {
    createParticles();
    
    // Add resize listener to adjust particles
    window.addEventListener('resize', function() {
      const particlesContainer = document.getElementById('particles');
      particlesContainer.innerHTML = '';
      createParticles();
    });
  });
  
  document.querySelector('.form-group').addEventListener('click', function() {
  document.getElementById('folderInput').click();
  });
  
  document.getElementById('chooseDirectoryBtn').addEventListener('click', function() {
  document.getElementById('folderInput').click();
  });
  
  document.getElementById('folderInput').addEventListener('change', function(e) {
  const files = e.target.files;
  
  const docxFiles = [];
  
  // Filter files but don't display in the folder upload area
  let docxCount = 0;
  for (let i = 0; i < files.length; i++) {
  if (files[i].name.endsWith('.docx')) {
    docxCount++;
    docxFiles.push(files[i]); // Store for main files section
    
  
  }
  }
  
  
  
  if (docxCount === 0) {
  showMessage('No .docx files found in the selected folder', 'error');
  } else {
  showMessage(`Found ${docxCount} .docx files`, 'info');
  
  // Keep this - display the files in the main files section only
  displayFilesInMainSection(docxFiles);
  
  // Enable clear button
  document.getElementById('clearBtn').disabled = false;
  }
  });
  
  document.getElementById('generateXMLBtn').addEventListener('click', function() {
  const files = document.getElementById('folderInput').files;
  if (files.length === 0) {
  showMessage('Please select a folder', 'error');
  return;
  }
  
  const docxFiles = Array.from(files).filter(file => file.name.endsWith('.docx'));
  if (docxFiles.length === 0) {
  showMessage('No .docx files found in the selected folder', 'error');
  return;
  }
  
  showMessage(`Generating XML for ${docxFiles.length} files, please wait...`, 'info');
  showProgress('folder', true);
  
  // Update status indicators in the main files list to show processing
  updateFileStatuses('processing');
  
  const formData = new FormData();
  for (let i = 0; i < docxFiles.length; i++) {
  formData.append('files', docxFiles[i]);
  updateProgress('folder', (i / docxFiles.length) * 50); // First half of progress
  }
  
  // Call your XML generation endpoint
  fetch('/api/generate-xml', {
  method: 'POST',
  body: formData
  })
  .then(response => {
  updateProgress('folder', 75); // 75% progress
  if (!response.ok) {
    return response.json().then(data => {
      // Update file statuses to failed
      updateFileStatuses('error');
      throw new Error(data.error || 'Failed to generate XML');
    });
  }
  return response.blob();
  })
  .then(blob => {
  updateProgress('folder', 100); // 100% progress
  // Update file statuses to processed
  updateFileStatuses('success');
  
  // Create a download link for the XML file
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'updatedWord2jld.xml'; // Match the backend's download_name
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
  
  showMessage('XML generation completed successfully! Download started.', 'success');
  setTimeout(() => hideProgress('folder'), 2000);
  })
  .catch(error => {
  showMessage('Error: ' + error.message, 'error');
  hideProgress('folder');
  });
  });
  
  // Function to display files in the main files section
  function displayFilesInMainSection(files) {
  const fileList = document.getElementById('fileList');
  
  // Clear the "No files uploaded yet" message
  fileList.innerHTML = '';
  
  // Update counters
  document.getElementById('totalFilesCount').textContent = files.length;
  document.getElementById('processedFilesCount').textContent = '0';
  document.getElementById('pendingFilesCount').textContent = files.length;
  document.getElementById('failedFilesCount').textContent = '0';
  
  // Create file items
  files.forEach(file => {
  // Create list item
  const li = document.createElement('li');
  li.className = 'file-item';
  li.dataset.status = 'pending';
  li.dataset.type = 'docx';
  li.dataset.path = file.webkitRelativePath || file.name;
  
  // Using your existing structure to maintain your CSS styling
  li.innerHTML = `
    <div class="file-info">
      <div class="file-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
          stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
          <line x1="16" y1="13" x2="8" y2="13"></line>
          <line x1="16" y1="17" x2="8" y2="17"></line>
          <polyline points="10 9 9 9 8 9"></polyline>
        </svg>
      </div>
      <div class="file-details">
        <div class="file-name">${file.name}</div>
        <div class="file-path">${file.webkitRelativePath ? file.webkitRelativePath.replace('/' + file.name, '') : 'Root'}</div>
      </div>
      <div class="file-size">${formatFileSize(file.size)}</div>
    </div>
    <div class="status-indicator pending">
      <span class="status-dot"></span>
      <span class="status-text">Pending</span>
    </div>
  `;
  
  // Add to file list
  fileList.appendChild(li);
  });
  }
  
  // Helper function to format file size
  function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  else return (bytes / 1048576).toFixed(1) + ' MB';
  }
  
  document.getElementById('folderForm').addEventListener('submit', function(e) {
  e.preventDefault();
  const files = document.getElementById('folderInput').files;
  if (files.length === 0) {
  showMessage('Please select a folder', 'error');
  return;
  }
  
  const docxFiles = Array.from(files).filter(file => file.name.endsWith('.docx'));
  if (docxFiles.length === 0) {
  showMessage('No .docx files found in the selected folder', 'error');
  return;
  }
  
  showMessage(`Processing ${docxFiles.length} files, please wait...`, 'info');
  showProgress('folder', true);
  
  // Update status indicators in the main files list
  updateFileStatuses('processing');
  
  const formData = new FormData();
  for (let i = 0; i < docxFiles.length; i++) {
  formData.append('files', docxFiles[i]);
  updateProgress('folder', (i / docxFiles.length) * 50); // First half of progress
  }
  
  fetch('/api/extract', {
  method: 'POST',
  body: formData
  })
  .then(response => {
    updateProgress('folder', 75); // 75% progress
    if (!response.ok) {
      return response.json().then(data => {
        // Update file statuses to failed
        updateFileStatuses('error');
        throw new Error(data.error || 'Failed to process files');
      });
    }
    return response.blob();
  })
  .then(blob => {
    updateProgress('folder', 100); // 100% progress
    // Update file statuses to processed
    updateFileStatuses('success');
    
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'StylesData.xlsx';
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    showMessage('Folder processed successfully! Download started.', 'success');
    setTimeout(() => hideProgress('folder'), 2000);
  })
  .catch(error => {
    showMessage('Error: ' + error.message, 'error');
    hideProgress('folder');
  });
  });
  
  // Function to update file statuses in the main files list
  function updateFileStatuses(status) {
  const fileItems = document.querySelectorAll('#fileList .file-item');
  
  // Count for updating the counters
  let total = fileItems.length;
  let processed = 0;
  let pending = 0;
  let failed = 0;
  
  fileItems.forEach(item => {
  // Update data attribute
  item.dataset.status = status;
  
  // Update status indicator
  const statusIndicator = item.querySelector('.status-indicator');
  statusIndicator.className = `status-indicator ${status}`;
  
  const statusText = item.querySelector('.status-text');
  
  // Set the appropriate text based on status
  if (status === 'processing') {
    statusText.textContent = 'Processing';
    pending = total;
  } else if (status === 'success') {
    statusText.textContent = 'Processed';
    processed = total;
  } else if (status === 'error') {
    statusText.textContent = 'Failed';
    failed = total;
  } else if (status === 'pending') {
    statusText.textContent = 'Pending';
    pending = total;
  }
  });
  
  // Update counters
  document.getElementById('totalFilesCount').textContent = total;
  document.getElementById('processedFilesCount').textContent = processed;
  document.getElementById('pendingFilesCount').textContent = pending;
  document.getElementById('failedFilesCount').textContent = failed;
  }
  
  // Clear button functionality
  document.getElementById('clearBtn').addEventListener('click', function() {
  const fileList = document.getElementById('fileList');
  
  // Clear the file list and show empty state
  fileList.innerHTML = `
  <div class="empty-state">
    <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1"
      stroke-linecap="round" stroke-linejoin="round"
      style="width: 80px; height: 80px; margin: 0 auto; color: #d9d9d9;">
      <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path>
      <polyline points="13 2 13 9 20 9"></polyline>
    </svg>
    <p>No files uploaded yet</p>
    <p style="font-size: 14px; margin-top: 8px;">Use the upload options above to start</p>
  </div>
  `;
  
  // Reset counters
  document.getElementById('totalFilesCount').textContent = '0';
  document.getElementById('processedFilesCount').textContent = '0';
  document.getElementById('pendingFilesCount').textContent = '0';
  document.getElementById('failedFilesCount').textContent = '0';
  
  // Disable clear button
  this.disabled = true;
  
  // Clear folder input
  // document.getElementById('folderInput').value = '';
  // document.getElementById('folderFilesList').innerHTML = '';
  
  // showMessage('Files cleared', 'info');
  document.getElementById('folderInput').value = '';
  // This line can stay since we'll still have the element in HTML, just empty
  document.getElementById('folderFilesList').innerHTML = '';
  
  showMessage('Files cleared', 'info');
  });
  
  // Add file filtering functionality
  document.getElementById('searchInput').addEventListener('input', filterFiles);
  document.getElementById('statusFilter').addEventListener('change', filterFiles);
  document.getElementById('typeFilter').addEventListener('change', filterFiles);
  
  function filterFiles() {
  const searchText = document.getElementById('searchInput').value.toLowerCase();
  const statusFilter = document.getElementById('statusFilter').value;
  const typeFilter = document.getElementById('typeFilter').value;
  
  const fileItems = document.querySelectorAll('#fileList .file-item');
  
  fileItems.forEach(item => {
  const fileName = item.querySelector('.file-name').textContent.toLowerCase();
  const filePath = item.querySelector('.file-path').textContent.toLowerCase();
  const fileStatus = item.dataset.status;
  const fileType = item.dataset.type;
  
  const matchesSearch = fileName.includes(searchText) || filePath.includes(searchText);
  const matchesStatus = statusFilter === 'all' || fileStatus === statusFilter;
  const matchesType = typeFilter === 'all' || fileType === typeFilter;
  
  if (matchesSearch && matchesStatus && matchesType) {
    item.style.display = '';
  } else {
    item.style.display = 'none';
  }
  });
  }
  
  function showMessage(text, type) {
  const messageEl = document.getElementById('message');
  messageEl.textContent = text;
  messageEl.className = type || 'info';
  }
  
  function showProgress(type, show) {
  const progressBar = document.getElementById(`${type}ProgressBar`);
  if (show) {
  progressBar.style.display = 'block';
  updateProgress(type, 0);
  } else {
  progressBar.style.display = 'none';
  }
  }
  
  function hideProgress(type) {
  const progressBar = document.getElementById(`${type}ProgressBar`);
  progressBar.style.display = 'none';
  }
  
  function updateProgress(type, percent) {
  document.getElementById(`${type}ProgressFill`).style.width = `${percent}%`;
  }
  
  
  
  
  const menuToggle = document.getElementById('menuToggle');
  const navContent = document.getElementById('navContent');
  
  menuToggle.addEventListener('click', () => {
    menuToggle.classList.toggle('open');
    navContent.classList.toggle('active');
  });
    
    const cards = document.querySelectorAll('.card');
  
    cards.forEach(card => {
      card.addEventListener('mousemove', e => {
        const rect = card.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
  
        card.style.setProperty('--x', `${x}px`);
        card.style.setProperty('--y', `${y}px`);
      });
  
      card.addEventListener('mouseleave', () => {
        // Optional: reset to center or keep last
        card.style.setProperty('--x', `50%`);
        card.style.setProperty('--y', `50%`);
      });
    });
  
  
    