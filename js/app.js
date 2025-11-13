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

    // Check if mammoth library is loaded
    if (typeof mammoth === 'undefined') {
        showError('Document parsing library not loaded. Please refresh the page and try again.');
        return;
    }

    try {
        console.log('Reading file:', file.name);
        const arrayBuffer = await file.arrayBuffer();
        console.log('File loaded, extracting text...');

        const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
        const text = result.value;
        console.log('Text extracted, length:', text.length);

        // Parse the extracted text into structured JSON
        resumeData = parseResumeText(text);
        console.log('Resume data parsed:', resumeData);

        displayPreview(resumeData);
        showSection('previewSection');
        hideSection('errorSection');
    } catch (error) {
        console.error('Error details:', error);
        showError('Error reading Word document: ' + error.message + '. Please make sure the file is a valid .docx file.');
    }
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

function displayPreview(data) {
    const preview = document.getElementById('resumePreview');
    let html = '';

    // Contact Information
    if (data.contact) {
        html += '<div class="contact-info">';
        html += `<h2>${data.contact.name || 'Name Not Provided'}</h2>`;
        if (data.contact.email) html += `<div class="contact-detail">📧 ${data.contact.email}</div>`;
        if (data.contact.phone) html += `<div class="contact-detail">📱 ${data.contact.phone}</div>`;
        if (data.contact.location) html += `<div class="contact-detail">📍 ${data.contact.location}</div>`;
        if (data.contact.linkedin) html += `<div class="contact-detail">🔗 ${data.contact.linkedin}</div>`;
        if (data.contact.github) html += `<div class="contact-detail">💻 ${data.contact.github}</div>`;
        html += '</div>';
    }

    // Objective
    if (data.objective) {
        html += '<h3>Objective</h3>';
        html += `<div class="objective">${data.objective}</div>`;
    }

    // Skills
    if (data.skills && data.skills.length > 0) {
        html += '<h3>Skills</h3>';
        html += '<div class="skills-list">';
        data.skills.forEach(skill => {
            html += `<span class="skill-tag">${skill}</span>`;
        });
        html += '</div>';
    }

    // Certificates
    if (data.certificates && data.certificates.length > 0) {
        html += '<h3>Certificates</h3>';
        html += '<div class="certificates-list">';
        data.certificates.forEach(cert => {
            html += `<div class="certificate-item">${cert}</div>`;
        });
        html += '</div>';
    }

    // Education
    if (data.education && data.education.length > 0) {
        html += '<h3>Education</h3>';
        html += '<div class="education-list">';
        data.education.forEach(edu => {
            html += '<div class="education-item">';
            html += `<h4>${edu.degree || 'Degree'}</h4>`;
            html += `<div class="item-meta">${edu.institution || ''}</div>`;
            if (edu.dates) html += `<div class="item-meta">${edu.dates}</div>`;
            if (edu.location) html += `<div class="item-meta">${edu.location}</div>`;
            if (edu.details && edu.details.length > 0) {
                edu.details.forEach(detail => {
                    html += `<div style="margin-top: 10px; color: #444;">${detail}</div>`;
                });
            }
            html += '</div>';
        });
        html += '</div>';
    }

    // Experience
    if (data.experience && data.experience.length > 0) {
        html += '<h3>Experience</h3>';
        html += '<div class="experience-list">';
        data.experience.forEach(exp => {
            html += '<div class="experience-item">';
            html += `<h4>${exp.title || 'Title'}</h4>`;
            html += `<div class="item-meta">${exp.company || ''}</div>`;
            if (exp.dates) html += `<div class="item-meta">${exp.dates}</div>`;
            if (exp.responsibilities && exp.responsibilities.length > 0) {
                html += '<ul>';
                exp.responsibilities.forEach(resp => {
                    html += `<li>${resp}</li>`;
                });
                html += '</ul>';
            }
            html += '</div>';
        });
        html += '</div>';
    }

    preview.innerHTML = html;
}

async function exportToWord() {
    if (!resumeData) {
        showError('No resume data loaded');
        return;
    }

    // Check if docx library is loaded
    if (typeof docx === 'undefined') {
        showError('Document generation library not loaded. Please refresh the page and try again.');
        return;
    }

    try {
        console.log('docx object:', docx);
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, UnderlineType } = docx;

        const children = [];

        // Contact Information
        if (resumeData.contact) {
            children.push(
                new Paragraph({
                    text: resumeData.contact.name || '',
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 }
                })
            );

            const contactDetails = [];
            if (resumeData.contact.phone) contactDetails.push(resumeData.contact.phone);
            if (resumeData.contact.email) contactDetails.push(resumeData.contact.email);
            if (resumeData.contact.location) contactDetails.push(resumeData.contact.location);
            if (resumeData.contact.linkedin) contactDetails.push(resumeData.contact.linkedin);
            if (resumeData.contact.github) contactDetails.push(resumeData.contact.github);

            if (contactDetails.length > 0) {
                children.push(
                    new Paragraph({
                        text: contactDetails.join(' | '),
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 }
                    })
                );
            }
        }

        // Objective
        if (resumeData.objective) {
            children.push(
                new Paragraph({
                    text: 'OBJECTIVE',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                })
            );
            children.push(
                new Paragraph({
                    text: resumeData.objective,
                    spacing: { after: 200 }
                })
            );
        }

        // Skills
        if (resumeData.skills && resumeData.skills.length > 0) {
            children.push(
                new Paragraph({
                    text: 'SKILLS',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                })
            );
            children.push(
                new Paragraph({
                    text: resumeData.skills.join(' • '),
                    spacing: { after: 200 }
                })
            );
        }

        // Certificates
        if (resumeData.certificates && resumeData.certificates.length > 0) {
            children.push(
                new Paragraph({
                    text: 'CERTIFICATES',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                })
            );
            resumeData.certificates.forEach(cert => {
                children.push(
                    new Paragraph({
                        text: '• ' + cert,
                        spacing: { after: 100 }
                    })
                );
            });
        }

        // Education
        if (resumeData.education && resumeData.education.length > 0) {
            children.push(
                new Paragraph({
                    text: 'EDUCATION',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                })
            );
            resumeData.education.forEach(edu => {
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: edu.institution || '',
                                bold: true
                            })
                        ],
                        spacing: { after: 50 }
                    })
                );
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: edu.degree || '',
                                italics: true
                            })
                        ],
                        spacing: { after: 50 }
                    })
                );
                if (edu.dates) {
                    children.push(
                        new Paragraph({
                            text: edu.dates,
                            spacing: { after: 50 }
                        })
                    );
                }
                if (edu.location) {
                    children.push(
                        new Paragraph({
                            text: edu.location,
                            spacing: { after: 50 }
                        })
                    );
                }
                if (edu.details && edu.details.length > 0) {
                    edu.details.forEach(detail => {
                        children.push(
                            new Paragraph({
                                text: detail,
                                spacing: { after: 50 }
                            })
                        );
                    });
                }
                children.push(
                    new Paragraph({
                        text: '',
                        spacing: { after: 100 }
                    })
                );
            });
        }

        // Experience
        if (resumeData.experience && resumeData.experience.length > 0) {
            children.push(
                new Paragraph({
                    text: 'EXPERIENCE',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 100 }
                })
            );
            resumeData.experience.forEach(exp => {
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: exp.company || '',
                                bold: true
                            })
                        ],
                        spacing: { after: 50 }
                    })
                );
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: exp.title || '',
                                italics: true
                            })
                        ],
                        spacing: { after: 50 }
                    })
                );
                if (exp.dates) {
                    children.push(
                        new Paragraph({
                            text: exp.dates,
                            spacing: { after: 100 }
                        })
                    );
                }
                if (exp.responsibilities && exp.responsibilities.length > 0) {
                    exp.responsibilities.forEach(resp => {
                        children.push(
                            new Paragraph({
                                text: '• ' + resp,
                                spacing: { after: 50 }
                            })
                        );
                    });
                }
                children.push(
                    new Paragraph({
                        text: '',
                        spacing: { after: 100 }
                    })
                );
            });
        }

        const doc = new Document({
            sections: [{
                properties: {},
                children: children
            }]
        });

        const blob = await Packer.toBlob(doc);
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
