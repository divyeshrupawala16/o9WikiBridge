let rules = null;
let htmlContent = "";

// HTML → Wiki Markup converter
function getRuleSyntax(rules, ruleName) {
  const rule = rules.rules.find(r => r.name === ruleName);
  return rule ? rule.syntax : null;
}

function convertHtmlToWiki(html, rules) {
  const headingSyntax = getRuleSyntax(rules, "Heading Levels");
  const tableSyntax = getRuleSyntax(rules, "Create a Table") || '{| class="wikitable"\n|Content goes in here\n|}';
  const tableStartSyntax = tableSyntax.start ? tableSyntax.start : "{|class='wikitable'";
  const tableEndSyntax = tableSyntax.end ? tableSyntax.end : "|}";
  const tableColSyntax = tableSyntax.column ? tableSyntax.column : "||";

  const boldSyntax = getRuleSyntax(rules, "Bold and Italics")?.bold || "'''";
  const italicSyntax = getRuleSyntax(rules, "Bold and Italics")?.italic || "''";
  const bulletSyntax = getRuleSyntax(rules, "Bullets") || "*";
  const numberingSyntax = getRuleSyntax(rules, "Numbering") || "#";
  const categorySyntax = getRuleSyntax(rules, "At the End of Each New Page Provide Category") || "[[Category: <category name>]]";
  const colorSyntax = getRuleSyntax(rules, "Color a Text") || '<span style="color:#000080">';
  const imageSyntax = getRuleSyntax(rules, "Images") || "";
  const imageStartSyntax = imageSyntax.start ? imageSyntax.start :  "{|class='wikitable'";
  const imageEndSyntax = imageSyntax.end ? imageSyntax.end : "|}";
  let imageIndex = 1;
  let isHeaderOpen = false;
  
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  function processInline(node) {
    let text = '';
    node.childNodes.forEach(child => {
      if (child.nodeType === 3) {
        text += child.textContent;
      } else if (child.nodeName === 'STRONG' || child.nodeName === 'B') {
        text += boldSyntax + child.textContent + boldSyntax;
      } else if (child.nodeName === 'EM' || child.nodeName === 'I') {
        text += italicSyntax + child.textContent + italicSyntax;
      } else if (child.nodeName === 'SPAN' && child.getAttribute('style')?.includes('color')) {
        text += `<span style="${child.getAttribute('style')}">${child.textContent}</span>`;
      } else if (child.nodeName === "IMG") {
        //let imageHeading = findNearestHeading(imageIndex, extractedImages.length);
        //imageHeading = imageHeading ? (imageHeading + "_"+imageIndex) : "NoImage"
        let imageHeading = "o9wiki-image-" + imageIndex;
        text+= imageStartSyntax;
        text += `|[[File:${imageHeading}.png]]`;    
        text+= imageEndSyntax;
        imageIndex = imageIndex +1;
      } else if (child.nodeName === 'A') {
        // External/internal link
        const href = child.getAttribute('href');
        if (href) {
          text += `[${href} ${child.textContent}]`;
        } else {
          text += child.textContent;
        }
      } else {
        text += processInline(child);
      }

      if (text && child.childNodes.length) {
        child.childNodes.forEach(grandChild => {
          if (grandChild.nodeName === "IMG") {           
            text += `[[File:User Workflow Icon.png]]`; 
            imageIndex = imageIndex +1;
          }
        });
      }
    });
    
    return text;
  }

function processList(node, level = 1) {
  let out = '';

  // Decide if we need bullets (UL) or numbers (OL)
  let isUL = node.nodeName === "UL";
  let bullet = isUL ? bulletSyntax : '';
  let numbering = 1;

  node.childNodes.forEach(li => {
    if (li.nodeName === "LI") {
      let prefix;
      if (isUL) {
        prefix = bullet.repeat(level);
      } else {
        // For OL, get 'value' attribute or use current numbering
        let value = li.getAttribute && li.getAttribute('value');
        prefix = ((value ? value : numbering) + '.').padStart(level + 2, ' ');
        numbering++;
      }
      out += prefix + processInline(li) + '\n';

      // Nested lists
      li.childNodes.forEach(child => {
        if (child.nodeName === "UL" || child.nodeName === "OL") {
          out += processList(child, level + 1);
        }
      });
    }
  });
  return out;
}

  function processTable(tableNode) {
    let out = tableStartSyntax;
    let rows = tableNode.querySelectorAll('tr');
    let isSingleRow = rows.length === 1;
    rows.forEach((row, i) => {
      let cells = row.children;
      if (i === 0 && !isSingleRow) { // Header row
        for (let cell of cells) {
          let text = processInline(cell);
          // Remove concatenated empty quotes at the end (one or more pairs of '')
          text = text.replace(/^(')+|(')+$/g, '');//text.replace(/('')+$/g, ''); 
          out += '!' + text + '\n';
        }
      } else {
        out += '|-\n';
        for (let cell of cells) {
          let text = processInline(cell);
          // Remove concatenated empty quotes at the end (one or more pairs of '')
          text = text.replace(/^(')+|(')+$/g, '');//text.replace(/('')+$/g, ''); 
          out += tableColSyntax + text + '\n';
        }
      }
    });

    if (out && (out.toLowerCase().indexOf("info") >-1 || out.toLowerCase().indexOf("note") >-1 || out.toLowerCase().indexOf("shortcut ") > -1 || out.toLowerCase().indexOf("collab ") > -1)) {
      out = out.replace(/User Workflow Icon\.png/g, 'Note Icon1.png');
    }
    
    out += tableEndSyntax;
    return out;
  }

  function processNode(node) {
    // Headings
    if (/^H[1-6]$/.test(node.nodeName)) {
      let level = parseInt(node.nodeName.charAt(1));
      let headingWiki = "";
      if (isHeaderOpen) {
        headingWiki += "\n|}\n\n";
      }
      headingWiki += '='.repeat(level) + node.textContent.trim() + '='.repeat(level);
      headingWiki += '\n\n {| class="wikitable"'
      isHeaderOpen = true;
      return headingWiki + '\n|\n\n';
    }
    // Paragraphs
    if (node.nodeName === "P") {
      return processInline(node) + '\n\n';
    }
    // Unordered/ordered lists
    if (node.nodeName === "UL" || node.nodeName === "OL") {
      return processList(node) + '\n';
    }
    // Tables
    if (node.nodeName === "TABLE") {
      return processTable(node) + '\n\n';
    }
    // Images
    if (node.nodeName === "IMG") {
      const src = node.getAttribute('src');
      const alt = node.getAttribute('alt') || '';
      return `[[File:${src}|${alt}]]\n`;
    }
    // Category (if present as a special marker)
    if (node.nodeName === "CATEGORY") {
      return categorySyntax + '\n';
    }
    // Otherwise, walk children
    let wiki = '';
    node.childNodes.forEach(child => {
      wiki += processNode(child);
    });
    return wiki;
  }

  var wikiMarkup = processNode(doc.body);
  wikiMarkup += "|}";
  return wikiMarkup.replace(/\n{3,}/g, '\n\n');
}

// New code
// Global variables
let wikiRules = null;
let currentTab = 'rules';

const docUploadSection = document.getElementById('docUploadSection');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const fileType = document.getElementById('fileType');
const originalContent = document.getElementById('originalContent');
const wikiOutput = document.getElementById('wikiOutput');
const copyBtn = document.getElementById('copyBtn');
const processing = document.getElementById('processing');
const successMessage = document.getElementById('successMessage');

// Tab switching
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        const tabName = tab.dataset.tab;
        switchTab(tabName);
    });
});

function switchTab(tabName) {
    // Update active tab
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    // Update active content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${tabName}-tab`).classList.add('active');

    currentTab = tabName;

    // Prevent conversion tab if no rules loaded
    if (tabName === 'convert' && !rules) {
        alert('Please upload rules first before converting documents.');
        switchTab('rules');
    }
}

const rulesInput = document.getElementById('rulesFileInput');
const rulesUploadSection = document.getElementById('rulesUploadSection');

rulesInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleRulesFile(e.target.files[0]);
    }
});

// Drag and drop for rules
rulesUploadSection.addEventListener('dragover', (e) => {
    e.preventDefault();
    rulesUploadSection.classList.add('dragover');
});

rulesUploadSection.addEventListener('dragleave', (e) => {
    e.preventDefault();
    rulesUploadSection.classList.remove('dragover');
});

rulesUploadSection.addEventListener('drop', (e) => {
    e.preventDefault();
    rulesUploadSection.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleRulesFile(files[0]);
    }
});

 function handleRulesFile(file) {
    if (!file.name.endsWith('.json')) {
        alert('Please upload a .JSON file with wiki rules.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const rulesText = e.target.result;
            rules = JSON.parse(e.target.result);
            //wikiRules = parseRulesFromText(rulesText);
            updateRulesStatus('loaded', file.name);
            showRulesPreview(rulesText);
        } catch (error) {
            alert('Error parsing rules file: ' + error.message);            
        }
    };
    reader.readAsText(file);
}

function updateRulesStatus(type, fileName = '') {
    const statusEl = document.getElementById('rulesStatus');
    const statusTextEl = document.getElementById('rulesStatusText');
    
    statusEl.className = 'rules-status';
    
    if (type === 'loaded') {
        statusEl.classList.add('rules-loaded');
        statusTextEl.textContent = `✅ Custom rules loaded from: ${fileName}`;
    } else if (type === 'default') {
        statusEl.classList.add('rules-default');
        statusTextEl.textContent = '📋 Using default wiki formatting rules';
    }
}

function showRulesPreview(customRulesText = '') {
    const previewEl = document.getElementById('rulesPreview');
    const contentEl = document.getElementById('rulesContent');
    
    if (customRulesText) {
        contentEl.textContent = customRulesText.substring(0, 1000) + (customRulesText.length > 1000 ? '...' : '');
    } else {
        contentEl.textContent = 'Default Wiki Rules Loaded:\n- Headings: ==Text==, ===Text===\n- Bold: \'\'\'Text\'\'\'\n- Tables: {| class="wikitable"\n- Lists: * Item, # Item\n- And more...';
    }
    
    previewEl.classList.add('show');
}

const docInput = document.getElementById('wordFileInput');
docInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleDocFile(e.target.files[0]);
    }
});


// Drag and drop for documents
docUploadSection.addEventListener('dragover', (e) => {
    e.preventDefault();
    docUploadSection.classList.add('dragover');
});

docUploadSection.addEventListener('dragleave', (e) => {
    e.preventDefault();
    docUploadSection.classList.remove('dragover');
});

docUploadSection.addEventListener('drop', (e) => {
    e.preventDefault();
    docUploadSection.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleDocFile(files[0]);
    }
});

function handleDocFile(file) {
    if (!rules) {
        alert('Please upload wiki rules first or use default rules.');
        switchTab('rules');
        return;
    }

    // Show file info
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileType.textContent = file.type || 'Unknown';
    fileInfo.classList.add('show');

    // Show processing
    processing.style.display = 'block';
    originalContent.value = '';
    wikiOutput.value = '';
    copyBtn.disabled = true;

    // Process file based on type
    if (file.name.endsWith('.docx')) {
      extractFromDOCX(file);

      mammoth.convertToHtml({arrayBuffer: file})
      .then(function(result) {
        htmlContent = result.value;
        //document.getElementById('docStatus').textContent = "✓ Document loaded";
        
        if (!rules || !htmlContent) return;
        const wikiMarkup = convertHtmlToWiki(htmlContent, rules);
        // document.getElementById('outputArea').value = wikiMarkup;
        // document.getElementById('copyBtn').disabled = false;

        wikiOutput.value = wikiMarkup;
        processing.style.display = 'none';
        copyBtn.disabled = false;
      });

      mammoth.extractRawText({arrayBuffer: file})
      .then(function(rawtext) {
        originalContent.value = rawtext.value;
      });

      
    } else {
        processTextFile(file);
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function copyToClipboard() {
    try {
        await navigator.clipboard.writeText(wikiOutput.value);
        successMessage.style.display = 'block';
        setTimeout(() => {
            successMessage.style.display = 'none';
        }, 3000);
    } catch (err) {
        // Fallback for older browsers
        wikiOutput.select();
        document.execCommand('copy');
        successMessage.style.display = 'block';
        setTimeout(() => {
            successMessage.style.display = 'none';
        }, 3000);
    }
}

// Download Images
const statusElem = document.getElementById('status');
const imagesGrid = document.getElementById('imagesGrid');
const downloadBtn = document.getElementById('downloadBtn');
let documentHeadings = [];
let extractedImages = [];

downloadBtn.addEventListener('click', downloadZip);


async function downloadZip() {
    if (extractedImages.length === 0) {
        showStatus('No images to download', 'error');
        return;
    }
    
    downloadBtn.disabled = true;
    downloadBtn.textContent = 'Creating ZIP...';
    
    try {
        const zip = new JSZip();
        
        for (const img of extractedImages) {
            zip.file(img.name, img.blob);
        }
        
        const content = await zip.generateAsync({type: 'blob'});
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = 'extracted_images.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showStatus('ZIP file downloaded successfully!', 'success');
        
    } catch (error) {
        showStatus('Error creating ZIP file: ' + error.message, 'error');
    }
    
    downloadBtn.disabled = false;
    downloadBtn.textContent = '📦 Download All Images as ZIP';
}

async function extractFromDOCX(file) {
  extractedImages = [];
  imagesGrid.innerHTML = '';
  imagesGrid.style.display = 'none';
  downloadBtn.style.display = 'none';

  try {
      // DOCX files are actually ZIP archives, so we can extract images directly
      const zip = await JSZip.loadAsync(file);
      const mediaFiles = [];

       // First, extract headings from document.xml
        try {
            const documentXml = await zip.file('word/document.xml').async('text');
            extractHeadingsFromDocumentXML(documentXml);
        } catch (xmlError) {
            console.warn('Could not extract headings from DOCX:', xmlError);
        }

      // Look for images in the word/media/ directory
      zip.forEach((relativePath, zipEntry) => {
          if (relativePath.startsWith('word/media/') && 
              (relativePath.endsWith('.jpg') || relativePath.endsWith('.jpeg') || 
                relativePath.endsWith('.png') || relativePath.endsWith('.gif') ||
                relativePath.endsWith('.bmp') || relativePath.endsWith('.tiff'))) {
              mediaFiles.push(zipEntry);
          }
      });
      
      if (mediaFiles.length === 0) {
          showStatus('No images found in DOCX file.', 'info');
          return;
      }
      
      for (let i = 0; i < mediaFiles.length; i++) {          
          const file = mediaFiles[i];
          const blob = await file.async('blob');
          const fileName = file.name.split('/').pop();
          
          extractedImages.push({
              name: fileName,
              blob: blob,
              url: URL.createObjectURL(blob)
          });
      }

      // Apply heading-based naming to all extracted images
      applyHeadingBasedNaming();

       if (extractedImages.length > 0) {
            showStatus(`Found ${extractedImages.length} images!`, 'success');
            displayImages();
            downloadBtn.style.display = 'block';
        } else {
            showStatus('No images found in the document.', 'info');
        }
      
  } catch (error) {
      throw new Error('Failed to extract images from DOCX: ' + error.message);
  }
}

 function showStatus(message, type = 'info') {
      statusElem.textContent = message;
      statusElem.className = `status ${type}`;
      statusElem.style.display = 'block';
}

function displayImages() {
    imagesGrid.innerHTML = '';
    
    extractedImages.forEach((img, index) => {
        const imageItem = document.createElement('div');
        imageItem.className = 'image-item';
        
        imageItem.innerHTML = `
            <img src="${img.url}" alt="Extracted image ${index + 1}" loading="lazy">
            <div class="image-name">${img.name}</div>
        `;
        
        imagesGrid.appendChild(imageItem);
    });
    
    imagesGrid.style.display = 'grid';
}

function extractHeadingsFromDocumentXML(xmlText) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
  
  // Find all paragraph elements
  const paragraphs = xmlDoc.getElementsByTagName('w:p');
  
  for (let para of paragraphs) {
      const styleElements = para.getElementsByTagName('w:pStyle');
      const textElements = para.getElementsByTagName('w:t');
      
      if (styleElements.length > 0 && textElements.length > 0) {
          const styleVal = styleElements[0].getAttribute('w:val');
          
          // Check if it's a heading style
          if (styleVal && (styleVal.toLowerCase().includes('heading') || 
                          styleVal.toLowerCase().includes('title') ||
                          /^h[1-6]$/i.test(styleVal))) {
              
              let headingText = '';
              for (let textEl of textElements) {
                  headingText += textEl.textContent || '';
              }
              
              if (headingText.trim()) {
                  const level = styleVal.match(/[1-6]/) ? parseInt(styleVal.match(/[1-6]/)[0]) : 1;
                  documentHeadings.push({
                      level: level,
                      text: headingText.trim()
                  });
              }
          }
      }
  }
}

function applyHeadingBasedNaming() {
    extractedImages.forEach((img, index) => {
        const headingPrefix = "o9wiki-image-" + (index+1);//findNearestHeading(index, extractedImages.length);
        const extension = img.name.split('.').pop();
        const newName = `${headingPrefix}.${extension}`; //`${headingPrefix}_${String(index + 1).padStart(2, '0')}.${extension}`;
        img.name = newName;
    });
}

function sanitizeFileName(text) {
    // Remove or replace invalid filename characters
    return text.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
              .replace(/\s+/g, '_')
              .substring(0, 50); // Limit length
}

function findNearestHeading(imageIndex, totalImages) {
    if (documentHeadings.length === 0) {
        return `Section_${Math.floor(imageIndex / Math.max(1, totalImages / 5)) + 1}`;
    }
    
    // Distribute images among available headings
    const headingIndex = Math.floor((imageIndex / totalImages) * documentHeadings.length);
    const heading = documentHeadings[Math.min(headingIndex, documentHeadings.length - 1)];
    
    return sanitizeFileName(heading.text);
}