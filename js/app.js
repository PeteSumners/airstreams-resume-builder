let resumeData = null;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    const jsonFileInput = document.getElementById('jsonFile');
    const docxFileInput = document.getElementById('docxFile');
    const exportBtn = document.getElementById('exportBtn');
    const exportJsonBtn = document.getElementById('exportJsonBtn');
    const loadNewBtn = document.getElementById('loadNewBtn');

    jsonFileInput.addEventListener('change', handleJsonFileSelect);
    docxFileInput.addEventListener('change', handleDocxFileSelect);
    exportBtn.addEventListener('click', exportToWord);
    exportJsonBtn.addEventListener('click', exportToJson);
    loadNewBtn.addEventListener('click', loadNewFile);
});

function handleJsonFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            resumeData = JSON.parse(e.target.result);
            displayPreview(resumeData);
            showSection('previewSection');
            hideSection('errorSection');
        } catch (error) {
            showError('Invalid JSON file: ' + error.message);
        }
    };
    reader.readAsText(file);
}

async function handleDocxFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Check if JSZip library is loaded
    if (typeof JSZip === 'undefined') {
        showError('Document parsing library not loaded. Please refresh the page and try again.');
        return;
    }

    try {
        console.log('Reading file:', file.name);
        const arrayBuffer = await file.arrayBuffer();
        console.log('File loaded, parsing docx structure...');

        // Parse the docx file using JSZip
        const zip = await JSZip.loadAsync(arrayBuffer);
        const documentXml = await zip.file('word/document.xml').async('text');

        console.log('Document XML loaded, parsing tables...');

        // Parse the XML to extract table data
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, 'text/xml');

        // Extract structured data from tables
        resumeData = parseDocxTables(xmlDoc);
        console.log('Resume data parsed:', resumeData);

        displayPreview(resumeData);
        showSection('previewSection');
        hideSection('errorSection');
    } catch (error) {
        console.error('Error details:', error);
        showError('Error reading Word document: ' + error.message + '. Please make sure the file is a valid .docx file.');
    }
}

function parseDocxTables(xmlDoc) {
    const data = {
        contact: {},
        objective: '',
        skills: [],
        certificates: [],
        education: [],
        experience: []
    };

    // Helper function to get text from a cell
    function getCellText(cell) {
        const textNodes = cell.getElementsByTagNameNS('*', 't');
        let text = '';
        for (let i = 0; i < textNodes.length; i++) {
            text += textNodes[i].textContent;
        }
        return text.trim();
    }

    // Helper function to get all paragraphs from a cell
    function getCellParagraphs(cell) {
        const paragraphs = [];
        const pNodes = cell.getElementsByTagNameNS('*', 'p');
        for (let i = 0; i < pNodes.length; i++) {
            const text = getCellText(pNodes[i]);
            if (text) paragraphs.push(text);
        }
        return paragraphs;
    }

    // Get all tables
    const tables = xmlDoc.getElementsByTagNameNS('*', 'tbl');
    console.log('Found', tables.length, 'tables');

    if (tables.length < 6) {
        console.warn('Expected at least 6 tables, found', tables.length);
    }

    // Table 0: Contact Information (2x2)
    if (tables[0]) {
        const rows = tables[0].getElementsByTagNameNS('*', 'tr');
        if (rows.length >= 2) {
            const row0Cells = rows[0].getElementsByTagNameNS('*', 'tc');
            const row1Cells = rows[1].getElementsByTagNameNS('*', 'tc');

            if (row0Cells.length >= 2) {
                data.contact.name = getCellText(row0Cells[0]);
                data.contact.phone = getCellText(row0Cells[1]);
            }
            if (row1Cells.length >= 2) {
                data.contact.location = getCellText(row1Cells[0]);
                data.contact.email = getCellText(row1Cells[1]);
            }
        }
    }

    // Table 1: Objective
    if (tables[1]) {
        const rows = tables[1].getElementsByTagNameNS('*', 'tr');
        if (rows.length > 0) {
            const cells = rows[0].getElementsByTagNameNS('*', 'tc');
            if (cells.length > 0) {
                data.objective = getCellText(cells[0]);
            }
        }
    }

    // Table 2: Skills (2 columns)
    if (tables[2]) {
        const rows = tables[2].getElementsByTagNameNS('*', 'tr');
        for (let i = 0; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagNameNS('*', 'tc');
            for (let j = 0; j < cells.length; j++) {
                const paragraphs = getCellParagraphs(cells[j]);
                paragraphs.forEach(p => {
                    if (p) data.skills.push(p);
                });
            }
        }
    }

    // Table 3: Certificates (2 columns)
    if (tables[3]) {
        const rows = tables[3].getElementsByTagNameNS('*', 'tr');
        for (let i = 0; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagNameNS('*', 'tc');
            for (let j = 0; j < cells.length; j++) {
                const text = getCellText(cells[j]);
                if (text) data.certificates.push(text);
            }
        }
    }

    // Table 4: Education (institution/date layout)
    if (tables[4]) {
        const rows = tables[4].getElementsByTagNameNS('*', 'tr');
        let i = 0;
        while (i < rows.length) {
            const cells = rows[i].getElementsByTagNameNS('*', 'tc');
            if (cells.length >= 2) {
                const institutionParagraphs = getCellParagraphs(cells[0]);
                const dates = getCellText(cells[1]);

                const eduEntry = {
                    institution: institutionParagraphs[0] || '',
                    degree: institutionParagraphs[1] || '',
                    dates: dates,
                    location: '',
                    details: []
                };

                // Check if next row has details
                if (i + 1 < rows.length) {
                    const nextCells = rows[i + 1].getElementsByTagNameNS('*', 'tc');
                    if (nextCells.length > 0) {
                        const detailText = getCellText(nextCells[0]);
                        if (detailText && !detailText.includes('educational institute')) {
                            eduEntry.details = [detailText];
                            i++; // Skip the detail row
                        }
                    }
                }

                if (eduEntry.institution || eduEntry.degree) {
                    data.education.push(eduEntry);
                }
            }
            i++;
        }
    }

    // Table 5: Experience (company/title/date + responsibilities)
    if (tables[5]) {
        const rows = tables[5].getElementsByTagNameNS('*', 'tr');
        let i = 0;
        while (i < rows.length) {
            const cells = rows[i].getElementsByTagNameNS('*', 'tc');
            if (cells.length >= 3) {
                const expEntry = {
                    company: getCellText(cells[0]),
                    title: getCellText(cells[1]),
                    dates: getCellText(cells[2]),
                    responsibilities: []
                };

                // Check if next row has responsibilities
                if (i + 1 < rows.length) {
                    const nextCells = rows[i + 1].getElementsByTagNameNS('*', 'tc');
                    if (nextCells.length > 0) {
                        // Get all paragraphs from first cell (they contain the bullet points)
                        const responsibilities = getCellParagraphs(nextCells[0]);
                        expEntry.responsibilities = responsibilities;
                        i++; // Skip the responsibility row
                    }
                }

                if (expEntry.company || expEntry.title) {
                    data.experience.push(expEntry);
                }
            }
            i++;
        }
    }

    return data;
}

function parseResumeText(text) {
    const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);

    const data = {
        contact: {},
        objective: '',
        skills: [],
        certificates: [],
        education: [],
        experience: []
    };

    let currentSection = null;
    let currentItem = null;

    // Common section headers
    const sectionHeaders = {
        'objective': ['objective', 'summary', 'professional summary'],
        'skills': ['skills', 'technical skills', 'core competencies'],
        'certificates': ['certificates', 'certifications', 'licenses'],
        'education': ['education', 'education and training'],
        'experience': ['experience', 'work experience', 'career history', 'employment history']
    };

    // Extract name (usually first line or one of the first few lines)
    if (lines.length > 0) {
        data.contact.name = lines[0];
    }

    // Look for email, phone, location in first few lines
    for (let i = 0; i < Math.min(10, lines.length); i++) {
        const line = lines[i];

        // Email detection
        const emailMatch = line.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
        if (emailMatch) {
            data.contact.email = emailMatch[1];
        }

        // Phone detection
        const phoneMatch = line.match(/(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})/);
        if (phoneMatch) {
            data.contact.phone = phoneMatch[1];
        }

        // Location detection (City, State pattern)
        const locationMatch = line.match(/([A-Z][a-z]+,\s*[A-Z]{2}|[A-Z][a-z]+,\s*[A-Z][a-z]+)/);
        if (locationMatch && !line.includes('@') && !phoneMatch) {
            data.contact.location = locationMatch[1];
        }
    }

    // Process the rest of the document
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const lowerLine = line.toLowerCase();

        // Check if this is a section header
        let foundSection = false;
        for (const [section, headers] of Object.entries(sectionHeaders)) {
            if (headers.some(header => lowerLine === header || lowerLine.startsWith(header))) {
                currentSection = section;
                foundSection = true;
                currentItem = null;
                break;
            }
        }

        if (foundSection) continue;

        // Process content based on current section
        if (currentSection === 'objective') {
            if (data.objective) {
                data.objective += ' ' + line;
            } else {
                data.objective = line;
            }
        } else if (currentSection === 'skills') {
            // Split by common delimiters
            const skills = line.split(/[•,|]/).map(s => s.trim()).filter(s => s.length > 0);
            data.skills.push(...skills);
        } else if (currentSection === 'certificates') {
            if (line.startsWith('•') || line.startsWith('-')) {
                data.certificates.push(line.replace(/^[•\-]\s*/, ''));
            } else {
                data.certificates.push(line);
            }
        } else if (currentSection === 'education') {
            // Try to detect new education entry (usually has dates or institution name)
            if (line.match(/\d{4}/) || (currentItem && line.length > 20)) {
                if (currentItem && currentItem.institution) {
                    data.education.push(currentItem);
                }
                currentItem = {
                    institution: line,
                    degree: '',
                    dates: '',
                    location: '',
                    details: []
                };
            } else if (currentItem) {
                if (!currentItem.degree) {
                    currentItem.degree = line;
                } else if (line.match(/\d{4}/)) {
                    currentItem.dates = line;
                } else {
                    currentItem.details.push(line);
                }
            }
        } else if (currentSection === 'experience') {
            // Try to detect new experience entry
            if (line.match(/\d{4}/) && line.includes('–') || line.includes('-') && line.match(/\d{4}/)) {
                // This looks like a date range
                if (currentItem && currentItem.company) {
                    data.experience.push(currentItem);
                }
                if (i > 0) {
                    currentItem = {
                        company: lines[i - 1],
                        title: '',
                        dates: line,
                        responsibilities: []
                    };
                }
            } else if (line.startsWith('•') || line.startsWith('-')) {
                if (currentItem) {
                    currentItem.responsibilities.push(line.replace(/^[•\-]\s*/, ''));
                }
            } else if (currentItem && !currentItem.title) {
                currentItem.title = line;
            } else if (!currentItem || (currentItem && currentItem.title && line.length > 20)) {
                // Might be a new company
                if (currentItem && currentItem.company) {
                    data.experience.push(currentItem);
                }
                currentItem = {
                    company: line,
                    title: '',
                    dates: '',
                    responsibilities: []
                };
            }
        }
    }

    // Push any remaining items
    if (currentSection === 'education' && currentItem && currentItem.institution) {
        data.education.push(currentItem);
    }
    if (currentSection === 'experience' && currentItem && currentItem.company) {
        data.experience.push(currentItem);
    }

    return data;
}

async function displayPreview(data) {
    const preview = document.getElementById('resumePreview');

    // Check if docx-preview is loaded (it adds renderAsync to the docx namespace)
    console.log('Checking for docx-preview...');
    console.log('docx.renderAsync:', typeof docx.renderAsync);

    if (typeof docx.renderAsync === 'function') {
        // Use docx-preview to render actual DOCX
        console.log('✓ Using true DOCX rendering with docx-preview...');
        try {
            const doc = await generateDocxDocument(data);
            // Use docxLib (original library) for Packer
            const docxLibrary = window.docxLib || window.docx;
            const blob = await docxLibrary.Packer.toBlob(doc);

            // Clear preview and render DOCX
            preview.innerHTML = '';
            preview.style.padding = '20px';
            preview.style.backgroundColor = '#f5f5f5';

            // Use docx.renderAsync from docx-preview library
            await docx.renderAsync(blob, preview, null, {
                className: 'docx-preview-container',
                inWrapper: false, // Don't wrap in container - let it expand naturally
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreLastRenderedPageBreak: false,
                experimental: false,
                trimXmlDeclaration: true,
                useBase64URL: false,
                useMathMLPolyfill: false,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true,
                debug: false
            });

            console.log('✓ DOCX preview rendered successfully!');
            return;
        } catch (error) {
            console.error('✗ Error rendering DOCX preview:', error);
            console.log('Falling back to HTML preview');
            // Fall back to HTML preview
        }
    } else {
        console.log('docx-preview not loaded, using HTML fallback');
    }

    // Fallback: HTML preview
    // Create initial page
    let html = '<div class="page"><div class="page-number">Page 1</div>';

    // Contact Information Table (2x2)
    if (data.contact) {
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 10px;">';
        html += '<tr>';
        html += `<td style="width: 50%; font-weight: bold; font-size: 16pt;">${data.contact.name || 'Your Name'}</td>`;
        html += `<td style="width: 50%; text-align: right; font-weight: bold; font-size: 16pt;">${data.contact.phone || ''}</td>`;
        html += '</tr>';
        html += '<tr>';
        html += `<td style="font-size: 11pt;">${data.contact.location || ''}</td>`;
        html += `<td style="text-align: right; font-size: 11pt;">${data.contact.email || ''}</td>`;
        html += '</tr>';
        html += '</table>';
    }

    // Objective
    if (data.objective) {
        html += '<div style="text-align: center; font-weight: bold; font-size: 14pt; margin: 15px 0;">Objective</div>';
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">';
        html += '<tr><td style="text-align: justify; padding: 5px 0;">';
        html += data.objective;
        html += '</td></tr></table>';
    }

    // Skills and Qualifications (2-column table)
    if (data.skills && data.skills.length > 0) {
        html += '<div style="text-align: center; font-weight: bold; font-size: 14pt; margin: 15px 0;">Skills and Qualifications</div>';
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">';
        html += '<tr>';
        html += '<td style="width: 50%; vertical-align: top; padding-right: 10px;">';
        // Left column - even indexed skills
        for (let i = 0; i < data.skills.length; i += 2) {
            if (data.skills[i]) {
                html += `<div style="margin: 3px 0;">• ${data.skills[i]}</div>`;
            }
        }
        html += '</td>';
        html += '<td style="width: 50%; vertical-align: top; padding-left: 10px;">';
        // Right column - odd indexed skills
        for (let i = 1; i < data.skills.length; i += 2) {
            if (data.skills[i]) {
                html += `<div style="margin: 3px 0;">• ${data.skills[i]}</div>`;
            }
        }
        html += '</td>';
        html += '</tr></table>';
    }

    // Certificates (2-column table)
    if (data.certificates && data.certificates.length > 0) {
        html += '<div style="text-align: center; font-weight: bold; font-size: 14pt; margin: 15px 0;">Certificates</div>';
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">';

        // Calculate rows needed
        const certRows = Math.ceil(data.certificates.length / 2);
        for (let row = 0; row < certRows; row++) {
            html += '<tr>';
            const leftCert = data.certificates[row * 2];
            const rightCert = data.certificates[row * 2 + 1];
            html += '<td style="width: 50%; vertical-align: top; padding-right: 10px;">';
            if (leftCert) html += `<div style="margin: 3px 0;">• ${leftCert}</div>`;
            html += '</td>';
            html += '<td style="width: 50%; vertical-align: top; padding-left: 10px;">';
            if (rightCert) html += `<div style="margin: 3px 0;">• ${rightCert}</div>`;
            html += '</td>';
            html += '</tr>';
        }
        html += '</table>';
    }

    // Education and Training
    if (data.education && data.education.length > 0) {
        html += '<div style="text-align: center; font-weight: bold; font-size: 14pt; margin: 15px 0;">Education and Training</div>';

        data.education.forEach(edu => {
            html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">';
            // Header row: Institution/Degree | Date
            html += '<tr>';
            html += '<td style="width: 70%; font-weight: bold; vertical-align: top;">';
            html += edu.institution || '';
            if (edu.degree) html += `<br>${edu.degree}`;
            html += '</td>';
            html += `<td style="width: 30%; text-align: right; font-weight: bold; vertical-align: top;">${edu.dates || ''}</td>`;
            html += '</tr>';
            // Details row
            if (edu.details && edu.details.length > 0) {
                html += '<tr><td colspan="2" style="padding-top: 5px;">';
                edu.details.forEach(detail => {
                    html += `<div>${detail}</div>`;
                });
                html += '</td></tr>';
            }
            html += '</table>';
        });
    }

    // Career History
    if (data.experience && data.experience.length > 0) {
        html += '<div style="text-align: center; font-weight: bold; font-size: 14pt; margin: 15px 0;">Career History</div>';

        data.experience.forEach(exp => {
            html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">';
            // Header row: Company | Title | Dates
            html += '<tr>';
            html += `<td style="width: 40%; font-weight: bold; vertical-align: top;">${exp.company || 'Company Name'}</td>`;
            html += `<td style="width: 30%; text-align: center; font-weight: bold; vertical-align: top;">${exp.title || 'Job Title'}</td>`;
            html += `<td style="width: 30%; text-align: right; font-weight: bold; vertical-align: top;">${exp.dates || ''}</td>`;
            html += '</tr>';
            // Responsibilities row (merged across all columns)
            if (exp.responsibilities && exp.responsibilities.length > 0) {
                html += '<tr><td colspan="3" style="padding-top: 5px;">';
                exp.responsibilities.forEach(resp => {
                    html += `<div style="margin: 3px 0; text-align: justify;">• ${resp}</div>`;
                });
                html += '</td></tr>';
            }
            html += '</table>';
        });
    }

    html += '</div>'; // Close page

    preview.innerHTML = html;

    // Don't split into pages - let content flow naturally
    // The HTML preview is just for quick reference
    // The actual exported DOCX will have proper page breaks
}

// Split content into multiple pages
function splitIntoPages() {
    const preview = document.getElementById('resumePreview');
    const page = preview.querySelector('.page');
    if (!page) return;

    // Get all content (excluding page number)
    const pageNumber = page.querySelector('.page-number');
    const children = Array.from(page.children).filter(child => !child.classList.contains('page-number'));

    // Create a measuring div
    const measuringDiv = document.createElement('div');
    measuringDiv.style.position = 'absolute';
    measuringDiv.style.visibility = 'hidden';
    measuringDiv.style.width = '6.5in'; // 8.5in - 2in (for 1in margins on each side)
    document.body.appendChild(measuringDiv);

    // Disable page splitting for now - let content flow naturally in single page
    // The actual DOCX rendering with docx-preview handles page breaks correctly
    // This HTML fallback doesn't need perfect pagination

    // Just keep everything on one page or use a very large page height
    const pageHeight = 999999; // Effectively infinite - no page breaks in HTML preview
    let currentPage = createNewPage(1);
    let currentHeight = 0;

    children.forEach(child => {
        // Add child to current page (no breaking)
        currentPage.appendChild(child.cloneNode(true));
    });

    // Add final page if it has content
    if (currentPage.children.length > 1) { // More than just page number
        preview.appendChild(currentPage);
    }

    // Clean up
    document.body.removeChild(measuringDiv);

    // Remove the original page if we created new pages
    if (preview.querySelectorAll('.page').length > 1) {
        page.remove();
    }
}

// Create a new page element
function createNewPage(pageNumber) {
    const page = document.createElement('div');
    page.className = 'page';

    const pageNum = document.createElement('div');
    pageNum.className = 'page-number';
    pageNum.textContent = `Page ${pageNumber}`;
    page.appendChild(pageNum);

    return page;
}

// Generate DOCX document (without exporting)
async function generateDocxDocument(data) {
    // Use docxLib (the original docx.js library, saved before docx-preview loaded)
    const docxLibrary = window.docxLib || window.docx;

    if (typeof docxLibrary === 'undefined') {
        throw new Error('docx library not loaded');
    }

    const {
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, VerticalAlign, WidthType, BorderStyle,
        convertInchesToTwip
    } = docxLibrary;

    const children = [];

    // Use data parameter instead of resumeData
    const resumeData = data;

        // Helper function to create section header
        function createSectionHeader(text) {
            return new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        bold: true,
                        size: 28, // 14pt
                        font: 'Calibri'
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 200 }
            });
        }

        // Contact Information Table (2x2)
        if (resumeData.contact) {
            const contactTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({
                                    children: [new TextRun({
                                        text: resumeData.contact.name || 'Your Name',
                                        bold: true,
                                        size: 32, // 16pt
                                        font: 'Calibri'
                                    })],
                                    alignment: AlignmentType.LEFT
                                })],
                                width: { size: 50, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({
                                    children: [new TextRun({
                                        text: resumeData.contact.phone || '',
                                        bold: true,
                                        size: 32, // 16pt
                                        font: 'Calibri'
                                    })],
                                    alignment: AlignmentType.RIGHT
                                })],
                                width: { size: 50, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({
                                    text: resumeData.contact.location || '',
                                    alignment: AlignmentType.LEFT
                                })]
                            }),
                            new TableCell({
                                children: [new Paragraph({
                                    text: resumeData.contact.email || '',
                                    alignment: AlignmentType.RIGHT
                                })]
                            })
                        ]
                    })
                ]
            });
            children.push(contactTable);
            children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
        }

        // Objective
        if (resumeData.objective) {
            children.push(createSectionHeader('Objective'));
            const objectiveTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE }
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({
                                    text: resumeData.objective,
                                    alignment: AlignmentType.JUSTIFIED
                                })]
                            })
                        ]
                    })
                ]
            });
            children.push(objectiveTable);
            children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
        }

        // Skills and Qualifications (2-column table)
        if (resumeData.skills && resumeData.skills.length > 0) {
            children.push(createSectionHeader('Skills and Qualifications'));

            // Split skills into two columns
            const skillRows = [];
            for (let i = 0; i < resumeData.skills.length; i += 2) {
                skillRows.push(new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({
                                text: resumeData.skills[i] || '',
                                alignment: AlignmentType.LEFT
                            })],
                            width: { size: 50, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({
                                text: resumeData.skills[i + 1] || '',
                                alignment: AlignmentType.LEFT
                            })],
                            width: { size: 50, type: WidthType.PERCENTAGE }
                        })
                    ]
                }));
            }

            const skillsTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                },
                rows: skillRows
            });
            children.push(skillsTable);
            children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
        }

        // Certificates (2-column table with bullets)
        if (resumeData.certificates && resumeData.certificates.length > 0) {
            children.push(createSectionHeader('Certificates'));

            // Split certificates into two columns with bullets
            const certRows = [];
            for (let i = 0; i < resumeData.certificates.length; i += 2) {
                certRows.push(new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({
                                text: resumeData.certificates[i] || '',
                                alignment: AlignmentType.LEFT,
                                bullet: {
                                    level: 0
                                }
                            })],
                            width: { size: 50, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({
                                text: resumeData.certificates[i + 1] || '',
                                alignment: AlignmentType.LEFT,
                                bullet: {
                                    level: 0
                                }
                            })],
                            width: { size: 50, type: WidthType.PERCENTAGE }
                        })
                    ]
                }));
            }

            const certsTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                },
                rows: certRows
            });
            children.push(certsTable);
            children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
        }

        // Education and Training (table with institution/date columns)
        if (resumeData.education && resumeData.education.length > 0) {
            children.push(createSectionHeader('Education and Training'));

            resumeData.education.forEach(edu => {
                const eduRows = [];

                // First row: Institution/Program and Date
                eduRows.push(new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({
                                children: [
                                    new TextRun({
                                        text: (edu.institution || '') + (edu.degree ? '\n' + edu.degree : ''),
                                        bold: true
                                    })
                                ]
                            })],
                            width: { size: 70, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({
                                children: [new TextRun({
                                    text: edu.dates || '',
                                    bold: true
                                })],
                                alignment: AlignmentType.RIGHT
                            })],
                            width: { size: 30, type: WidthType.PERCENTAGE }
                        })
                    ]
                }));

                // Details rows (if any)
                if (edu.details && edu.details.length > 0) {
                    edu.details.forEach(detail => {
                        eduRows.push(new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({ text: detail })],
                                    columnSpan: 2
                                })
                            ]
                        }));
                    });
                }

                const eduTable = new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.NONE },
                        bottom: { style: BorderStyle.NONE },
                        left: { style: BorderStyle.NONE },
                        right: { style: BorderStyle.NONE },
                        insideHorizontal: { style: BorderStyle.NONE },
                        insideVertical: { style: BorderStyle.NONE }
                    },
                    rows: eduRows
                });
                children.push(eduTable);
                children.push(new Paragraph({ text: '', spacing: { after: 100 } }));
            });
        }

        // Career History (3-column table for experiences)
        if (resumeData.experience && resumeData.experience.length > 0) {
            children.push(createSectionHeader('Career History'));

            resumeData.experience.forEach(exp => {
                const expRows = [];

                // Header row: Company, Title, Dates
                expRows.push(new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({
                                children: [new TextRun({
                                    text: exp.company || 'Company Name',
                                    bold: true
                                })]
                            })],
                            width: { size: 40, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({
                                children: [new TextRun({
                                    text: exp.title || 'Job Title',
                                    bold: true
                                })]
                            })],
                            width: { size: 30, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({
                                children: [new TextRun({
                                    text: exp.dates || '',
                                    bold: true
                                })],
                                alignment: AlignmentType.RIGHT
                            })],
                            width: { size: 30, type: WidthType.PERCENTAGE }
                        })
                    ]
                }));

                // Responsibilities row (single merged cell with bullets)
                if (exp.responsibilities && exp.responsibilities.length > 0) {
                    // Create bullet paragraphs for each responsibility
                    const bulletParagraphs = exp.responsibilities.map(resp =>
                        new Paragraph({
                            text: resp,
                            alignment: AlignmentType.JUSTIFIED,
                            bullet: {
                                level: 0
                            }
                        })
                    );

                    // Add row with merged cell spanning all 3 columns
                    expRows.push(new TableRow({
                        children: [
                            new TableCell({
                                children: bulletParagraphs,
                                columnSpan: 3,
                                width: { size: 100, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }));
                }

                const expTable = new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.NONE },
                        bottom: { style: BorderStyle.NONE },
                        left: { style: BorderStyle.NONE },
                        right: { style: BorderStyle.NONE },
                        insideHorizontal: { style: BorderStyle.NONE },
                        insideVertical: { style: BorderStyle.NONE }
                    },
                    rows: expRows
                });
                children.push(expTable);
                children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
            });
        }

    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: 'Times New Roman',
                        size: 22 // 11pt
                    },
                    paragraph: {
                        spacing: {
                            line: 276, // 1.15 line spacing
                            before: 0,
                            after: 0
                        }
                    }
                }
            },
            paragraphStyles: [
                {
                    id: 'Normal',
                    name: 'Normal',
                    basedOn: 'Normal',
                    next: 'Normal',
                    run: {
                        font: 'Times New Roman',
                        size: 22
                    },
                    paragraph: {
                        spacing: {
                            line: 276,
                            before: 0,
                            after: 0
                        }
                    }
                }
            ]
        },
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: 1440, // 1 inch in twips
                        right: 1440,
                        bottom: 1440,
                        left: 1440
                    }
                }
            },
            children: children
        }]
    });

    return doc;
}

// Export to Word (uses generateDocxDocument)
async function exportToWord() {
    if (!resumeData) {
        showError('No resume data loaded');
        return;
    }

    try {
        const doc = await generateDocxDocument(resumeData);
        const docxLibrary = window.docxLib || window.docx;
        const blob = await docxLibrary.Packer.toBlob(doc);
        const fileName = (resumeData.contact && resumeData.contact.name)
            ? `${resumeData.contact.name.replace(/\s+/g, '_')}_Resume.docx`
            : 'Resume.docx';

        saveAs(blob, fileName);
    } catch (error) {
        console.error('Error generating document:', error);
        showError('Error generating Word document: ' + error.message);
    }
}

function exportToJson() {
    if (!resumeData) {
        showError('No resume data loaded');
        return;
    }

    try {
        const jsonString = JSON.stringify(resumeData, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json' });
        const fileName = (resumeData.contact && resumeData.contact.name)
            ? `${resumeData.contact.name.replace(/\s+/g, '_')}_resume.json`
            : 'resume.json';

        saveAs(blob, fileName);
    } catch (error) {
        showError('Error exporting JSON: ' + error.message);
    }
}

function loadNewFile() {
    document.getElementById('jsonFile').value = '';
    document.getElementById('docxFile').value = '';
    resumeData = null;
    hideSection('previewSection');
    hideSection('errorSection');
}

function showSection(sectionId) {
    document.getElementById(sectionId).style.display = 'block';
}

function hideSection(sectionId) {
    document.getElementById(sectionId).style.display = 'none';
}

function showError(message) {
    document.getElementById('errorMessage').textContent = message;
    showSection('errorSection');
    hideSection('previewSection');
}
