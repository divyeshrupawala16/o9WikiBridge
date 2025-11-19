let rules = null;
let htmlContent = "";

let imageRegistry = [];
let docNameGlobal = '';

// ============================================
// STEP 1: Pre-process HTML to add data-image-id to all <img> tags
// ============================================
function preprocessHtmlWithImageIds(html, docFileName) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  
  // Find ALL <img> tags in document order
  const allImages = doc.querySelectorAll('img');
  
  console.log(`Found ${allImages.length} images in HTML`);
  
  // Add sequential data-image-id attribute to each
  allImages.forEach((img, index) => {
    const imageId = index + 1; // 1-based indexing
    img.setAttribute('data-image-id', imageId);
    
    // Also add the standardized filename as data attribute for easy access
    const extension = 'png'; // Default to png for mammoth-generated images
    const standardName = `${docFileName}-${imageId}.${extension}`;
    img.setAttribute('data-image-name', standardName);
    
    console.log(`Tagged image ${imageId} as ${standardName}`);
  });
  
  return doc.documentElement.outerHTML;
}

// HTML → Wiki Markup converter
function getRuleSyntax(rules, ruleName) {
  const rule = rules.rules.find(r => r.name === ruleName);
  return rule ? rule.syntax : null;
}

function convertHtmlToWiki(html, rules, docFileName) {  
  const tableSyntax = getRuleSyntax(rules, "Create a Table");
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

  function processInline(node, isTriggeredFromTable = false) {
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
        // ⭐ Get the image ID from data attribute
        const imageId = child.getAttribute('data-image-id');
        const imageName = child.getAttribute('data-image-name');
        
        if (imageId && imageName) {
          console.log(`Processing image ID ${imageId}: ${imageName}`);
          
          if (isTriggeredFromTable) {
            text += `[[File:${imageName}]]`; 
          } else {          
            text += '\n\n' + tableStartSyntax + '\n |- \n | ';
            text += `[[File:${imageName}]]`; 
            text += "\n" + tableEndSyntax + "\n\n";
          }
        } else {
          console.warn('Image found without ID:', child);
          // Fallback: try to find in registry by any means
          text += '[[File:unknown-image.png]]';
        }
      } else if (child.nodeName === 'A') {
        const href = child.getAttribute('href');
        if (href) {
          text += `[${href} ${child.textContent}]`;
        } else {
          text += child.textContent;
        }
      } else if (child.nodeName === "UL" || child.nodeName === "OL") {
        text += processList(child) + '\n';
      } else {
        text += processInline(child, isTriggeredFromTable);
      }

      // Handle nested images (grandchildren)
      if (text && text.indexOf('File:') === -1 && child.childNodes.length) {
        child.childNodes.forEach(grandChild => {
          if (grandChild.nodeName === "IMG") {
            const imageId = grandChild.getAttribute('data-image-id');
            const imageName = grandChild.getAttribute('data-image-name');
            
            if (imageId && imageName) {
              console.log(`Processing nested image ID ${imageId}: ${imageName}`);
              text += `[[File:${imageName}]]`;
            }
          }
        });
      }
    });
    
    return text;
  }

  function processInlineForList(node) {
  let text = '';
  
  if (!node) return text;
  
  if (node.nodeType === 3) {
    return node.textContent;
  }
  
  // Get syntax from rules or use defaults
  const boldSyntax = rules ? (getRuleSyntax(rules, "Bold and Italics")?.bold || "'''") : "'''";
  const italicSyntax = rules ? (getRuleSyntax(rules, "Bold and Italics")?.italic || "''") : "''";
  
  if (node.nodeName === 'STRONG' || node.nodeName === 'B') {
    text += boldSyntax + node.textContent + boldSyntax;
  } else if (node.nodeName === 'EM' || node.nodeName === 'I') {
    text += italicSyntax + node.textContent + italicSyntax;
  } else if (node.nodeName === 'SPAN' && node.getAttribute && node.getAttribute('style')?.includes('color')) {
    text += `<span style="${node.getAttribute('style')}">${node.textContent}</span>`;
  } else if (node.nodeName === 'A') {
    const href = node.getAttribute && node.getAttribute('href');
    if (href) {
      text += `[${href} ${node.textContent}]`;
    } else {
      text += node.textContent;
    }
  } else if (node.nodeName !== 'UL' && node.nodeName !== 'OL') {
    // Process child nodes recursively
    if (node.childNodes) {
      node.childNodes.forEach(child => {
        text += processInlineForList(child);
      });
    }
  }
  
  return text;
}

// STEP 2: Replace your existing processList function with this:

function processList(node, level = 1) {
  let out = '';

  // Decide if we need bullets (UL) or numbers (OL)
  let isUL = node.nodeName === "UL";
  const bulletSyntax = getRuleSyntax(rules, "Bullets") || "*";
  let numbering = 1;
  let useLetters = false;

  node.childNodes.forEach(li => {
    if (li.nodeName === "LI") {
      // First, collect only text content (exclude nested lists)
      let textContent = '';
      li.childNodes.forEach(child => {
        if (child.nodeType === 3) { // Text node
          textContent += child.textContent;
        } else if (child.nodeName !== 'UL' && child.nodeName !== 'OL') {
          // Process inline elements but not nested lists
          textContent += processInlineForList(child);
        }
      });

      // Only add prefix if there's actual text content
      if (textContent.trim()) {
        let prefix;
        if (isUL) {
          prefix = bulletSyntax.repeat(1) + ' ';
        } else {
          // For OL, check if we should use letters based on level or list-style-type
          useLetters = false;
          
          // Check if parent OL has list-style-type attribute
          const listStyle = node.getAttribute && node.getAttribute('style');
          if (listStyle && listStyle.includes('list-style-type')) {
            useLetters = listStyle.includes('lower-alpha') || 
                         listStyle.includes('upper-alpha') ||
                         listStyle.includes('lower-latin') ||
                         listStyle.includes('upper-latin');
          }
          
          // Alternative: Use letters for nested lists (level > 1)
          if (level > 1 && !listStyle) {
            useLetters = true;
          }

          if (useLetters) {
            // Convert number to letter (1=a, 2=b, etc.)
            let value = li.getAttribute && li.getAttribute('value');
            let letterIndex = (value ? parseInt(value) : numbering) - 1;
            let letter = String.fromCharCode(97 + (letterIndex % 26)); // 97 is 'a'
            prefix = ''.repeat(level * 2) + letter + '. ';
          } else {
            // Use numbers for top-level or when explicitly numeric
            let value = li.getAttribute && li.getAttribute('value');
            let num = value ? value : numbering;
            prefix = ''.repeat((level - 1) * 2) + num + '. ';
          }
          numbering++;
        }
        
        if (useLetters) {
          out += ":";  
        }
        out += prefix + textContent.trim() + '<br>\n';       
      }

      // Handle nested lists separately
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
    let out = "";
    if (isHeaderOpen) {
      //out = "\n" + tableEndSyntax + "\n\n";
    }
    out += tableStartSyntax;
    let rows = tableNode.querySelectorAll('tr');
    let isSingleRow = rows.length === 1;
    rows.forEach((row, i) => {
      let cells = row.children;
      if (i === 0 && !isSingleRow) { // Header row
        for (let cell of cells) {
          let text = processInline(cell, true);
          // Remove concatenated empty quotes at the end (one or more pairs of '')
          text = text.replace(/^(')+|(')+$/g, '');//text.replace(/('')+$/g, ''); 
          out += '!' + text + '\n';
        }
      } else {
        out += '|-\n';
        for (let cell of cells) {          
          if (cell.textContent && cell.textContent == "-") {
            cell.textContent = "";            
          }
          let text = processInline(cell, true);
          // Remove concatenated empty quotes at the end (one or more pairs of '')
          text = text.replace(/^(')+|(')+$/g, '');//text.replace(/('')+$/g, ''); 
          let colSept = text.startsWith(bulletSyntax) || text.startsWith(numberingSyntax) ? "\n" : " ";
          out += tableColSyntax + colSept + text + '\n';
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
        // use configured table end syntax (was hardcoded '|}')
        headingWiki += "\n" + tableEndSyntax + "\n\n";
      }
      headingWiki += '='.repeat(level) + node.textContent.trim() + '='.repeat(level);
      // use configured table start syntax instead of hardcoded '{| class="wikitable"'
      headingWiki += '\n\n' + tableStartSyntax;
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

  // Remove empty wikitable blocks created when headings open a table but no rows were added.
  // Uses the configured tableStartSyntax / tableEndSyntax to find table blocks,
  // and deletes those whose inner content contains only whitespace or pipe characters.
  try {
    const startEsc = tableStartSyntax.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const endEsc = tableEndSyntax.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const emptyTableRegex = new RegExp(startEsc + '[\\s\\S]*?' + endEsc, 'g');

    wikiMarkup = wikiMarkup.replace(emptyTableRegex, match => {
      const inner = match.slice(tableStartSyntax.length, match.length - tableEndSyntax.length);
      // if inner contains only whitespace or pipe characters (no real rows/content), remove whole block
      if (/^[\s\|]*$/.test(inner)) return '';
      return match;
    });
  } catch (e) {
    // if anything fails, fall back to original markup
    console.warn('Empty table cleanup failed:', e);
  }
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

// Add this helper function to skip pages
function skipFirstPages(html, pagesToSkip) {
  if (pagesToSkip <= 0) return html;
  
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  
  // Look for page breaks (Word usually uses <hr> or specific page break markers)
  const pageBreaks = doc.querySelectorAll('hr, br[style*="page-break"], div[style*="page-break"]');
  
  if (pageBreaks.length >= pagesToSkip) {
    // Remove everything before the Nth page break
    const targetBreak = pageBreaks[pagesToSkip - 1];
    let currentNode = doc.body.firstChild;
    
    while (currentNode && currentNode !== targetBreak) {
      const nextNode = currentNode.nextSibling;
      currentNode.remove();
      currentNode = nextNode;
    }
    
    // Also remove the page break itself
    if (targetBreak) targetBreak.remove();
  } else {
    // If page breaks not found, try removing first N headings/sections
    const headings = doc.querySelectorAll('h1, h2, h3');
    if (headings.length > pagesToSkip) {
      for (let i = 0; i < pagesToSkip && i < headings.length; i++) {
        const heading = headings[i];
        // Remove all siblings until next heading
        let currentNode = heading;
        while (currentNode) {
          const nextNode = currentNode.nextSibling;
          if (nextNode && nextNode.nodeName && /^H[1-6]$/.test(nextNode.nodeName)) {
            break; // Stop at next heading
          }
          currentNode.remove();
          currentNode = nextNode;
        }
      }
    }
  }
  
  return doc.body.innerHTML;
}

function handleDocFile(file) {
    if (!rules) {
        alert('Please upload wiki rules first or use default rules.');
        switchTab('rules');
        return;
    }

    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileType.textContent = file.type || 'Unknown';
    fileInfo.classList.add('show');

    processing.style.display = 'block';
    originalContent.value = '';
    wikiOutput.value = '';
    copyBtn.disabled = true;

    if (file.name.endsWith('.docx')) {
      // First extract images to build registry
      extractFromDOCX(file);

      mammoth.convertToHtml({arrayBuffer: file})
      .then(function(result) {
        htmlContent = result.value;
        
        // Skip pages
        htmlContent = skipFirstPages(htmlContent, 6);
        htmlContent = fixOrderedListNumbering(htmlContent);

        if (!rules || !htmlContent) return;
        
        const docNameWithoutExt = file.name.split('.').slice(0, -1).join('.');
        
        // ⭐ CRITICAL: Preprocess HTML to add image IDs
        htmlContent = preprocessHtmlWithImageIds(htmlContent, docNameWithoutExt);
        
        console.log('Preprocessed HTML with image IDs');
        
        const wikiMarkup = convertHtmlToWiki(htmlContent, rules, docNameWithoutExt);

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

function fixOrderedListNumbering(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    let globalCounter = 1;
    let lastWasTable = false;
    
    // Find all elements in body
    const allElements = doc.body.querySelectorAll('*');
    
    allElements.forEach(element => {
        // Reset counter when we encounter a TABLE
        if (element.nodeName === 'TABLE') {
            globalCounter = 1;
            lastWasTable = true;
            return;
        }
        
        if (element.nodeName === 'OL') {
            // Check if this OL is inside a table
            const isInsideTable = element.closest('table') !== null;
            
            // If we just saw a table or we're inside a table, reset
            if (lastWasTable || isInsideTable) {
                globalCounter = 1;
            }
            
            lastWasTable = false;
            
            // Get all direct LI children
            const listItems = Array.from(element.children).filter(child => child.nodeName === 'LI');
            
            listItems.forEach((li, index) => {
                // Check if this LI already has a value attribute from Word
                const existingValue = li.getAttribute('value');
                
                if (existingValue) {
                    // Use the value from Word document
                    globalCounter = parseInt(existingValue);
                } else if (index === 0 && element.previousElementSibling) {
                    // This is the first item in a new OL
                    // Look backwards for the last OL before this one
                    let searchNode = element.previousElementSibling;
                    
                    while (searchNode && searchNode.nodeName !== 'OL') {
                        // If we hit a table, reset
                        if (searchNode.nodeName === 'TABLE') {
                            globalCounter = 1;
                            break;
                        }
                        searchNode = searchNode.previousElementSibling;
                    }
                    
                    if (searchNode && searchNode.nodeName === 'OL') {
                        // Found a previous OL, continue numbering
                        const prevItems = searchNode.querySelectorAll('li');
                        if (prevItems.length > 0) {
                            const lastItem = prevItems[prevItems.length - 1];
                            const lastValue = lastItem.getAttribute('value') || lastItem.getAttribute('data-counter');
                            globalCounter = lastValue ? parseInt(lastValue) + 1 : globalCounter;
                        }
                    } else {
                        // No previous OL found, reset to 1
                        globalCounter = 1;
                    }
                }
                
                // Set the value attribute
                li.setAttribute('value', globalCounter);
                li.setAttribute('data-counter', globalCounter);
                
                globalCounter++;
            });
        } else if (element.nodeName !== 'TABLE') {
            lastWasTable = false;
        }
    });
    
    return doc.body.innerHTML;
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

// NEW FUNCTION: Extract image order from document.xml
function extractImageOrderFromDocument(xmlText) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
    const imageRefs = [];
    
    // Look for drawing elements that reference images
    const drawings = xmlDoc.getElementsByTagName('a:blip');
    for (let drawing of drawings) {
        const embed = drawing.getAttribute('r:embed');
        if (embed) {
            imageRefs.push(embed);
        }
    }
    
    // Also check for v:imagedata (older format)
    const imageDatas = xmlDoc.getElementsByTagName('v:imagedata');
    for (let imageData of imageDatas) {
        const rId = imageData.getAttribute('r:id');
        if (rId) {
            imageRefs.push(rId);
        }
    }
    
    return imageRefs;
}

// NEW FUNCTION: Parse relationship file to map rId to media files
function parseRelationships(relsXml) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(relsXml, 'text/xml');
    const relationshipMap = {};
    
    const relationships = xmlDoc.getElementsByTagName('Relationship');
    for (let rel of relationships) {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        const type = rel.getAttribute('Type');
        
        // Check if it's an image relationship
        if (type && type.includes('image') && target) {
            relationshipMap[id] = target;
        }
    }
    
    return relationshipMap;
}

async function extractFromDOCX(file) {
  extractedImages = [];
  imageRegistry = [];
  imagesGrid.innerHTML = '';
  imagesGrid.style.display = 'none';
  downloadBtn.style.display = 'none';

  try {
      const zip = await JSZip.loadAsync(file);
      const mediaFiles = [];

      // Extract headings
      try {
          const documentXml = await zip.file('word/document.xml').async('text');
          extractHeadingsFromDocumentXML(documentXml);
      } catch (xmlError) {
          console.warn('Could not extract headings:', xmlError);
      }

      // Get image order from document
      let orderedImageRefs = [];
      try {
          const documentXml = await zip.file('word/document.xml').async('text');
          orderedImageRefs = extractImageOrderFromDocument(documentXml);
      } catch (error) {
          console.warn('Could not extract image order:', error);
      }

      // Parse relationships
      let relationshipMap = {};
      try {
          const relsXml = await zip.file('word/_rels/document.xml.rels').async('text');
          relationshipMap = parseRelationships(relsXml);
      } catch (error) {
          console.warn('Could not parse relationships:', error);
      }

      // Collect media files
      const mediaFileMap = {};
      zip.forEach((relativePath, zipEntry) => {
          if (relativePath.startsWith('word/media/') && 
              /\.(jpg|jpeg|png|gif|bmp|tiff)$/i.test(relativePath)) {
              const fileName = relativePath.split('/').pop();
              mediaFileMap[fileName] = zipEntry;
          }
      });
      
      // Process in document order
      if (orderedImageRefs.length > 0 && Object.keys(relationshipMap).length > 0) {
          for (const rId of orderedImageRefs) {
              const mediaPath = relationshipMap[rId];
              if (mediaPath) {
                  const fileName = mediaPath.split('/').pop();
                  const zipEntry = mediaFileMap[fileName];
                  if (zipEntry) {
                      mediaFiles.push(zipEntry);
                  }
              }
          }
      } else {
          const sortedEntries = Object.entries(mediaFileMap)
              .sort(([nameA], [nameB]) => {
                  const numA = parseInt(nameA.match(/\d+/) || [0]);
                  const numB = parseInt(nameB.match(/\d+/) || [0]);
                  return numA - numB;
              })
              .map(([_, entry]) => entry);
          mediaFiles.push(...sortedEntries);
      }

      if (mediaFiles.length === 0) {
          showStatus('No images found in DOCX file.', 'info');
          return;
      }
      
      // Get document name without extension
      const docNameWithoutExt = file.name.split('.').slice(0, -1).join('.');
      docNameGlobal = docNameWithoutExt; // Store globally
      
      // Process each image and build registry
      for (let i = 0; i < mediaFiles.length; i++) {          
          const zipFile = mediaFiles[i];
          const blob = await zipFile.async('blob');
          const originalFileName = zipFile.name.split('/').pop();
          const extension = originalFileName.split('.').pop();
          
          // Create standardized name matching the HTML preprocessing
          const standardizedName = `${docNameWithoutExt}-${i + 1}.${extension}`;
          
          const imageInfo = {
              id: i + 1,
              name: standardizedName,
              originalName: originalFileName,
              blob: blob,
              url: URL.createObjectURL(blob)
          };
          
          extractedImages.push(imageInfo);
          imageRegistry[i + 1] = imageInfo; // Store by ID for easy lookup
      }

      console.log('Image Registry created:', imageRegistry);

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

function applyHeadingBasedNaming(docFileName) {
  // Get document name without extension
    const docNameWithoutExt = docFileName.split('.').slice(0, -1).join('.');

    extractedImages.forEach((img, index) => {
        const headingPrefix = docNameWithoutExt+"-" + (index+1);//findNearestHeading(index, extractedImages.length);
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