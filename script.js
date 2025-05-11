let excelData = null;
let images = new Map();
let defaultMaleImage = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
let defaultFemaleImage = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
const LOGO_BASE64 = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
let uploadedLogoBase64 = '';

document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
document.getElementById('imageFiles').addEventListener('change', handleImageUpload);
document.getElementById('generateBtn').addEventListener('click', generatePDF);
document.getElementById('logoFile').addEventListener('change', handleLogoUpload);

function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('excelInfo').textContent = file.name;
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                excelData = XLSX.utils.sheet_to_json(firstSheet);
                // Remap data using real headers from the third row
                const headerRow = excelData.find(row => Object.values(row).includes('NOM & PRENOM'));
                if (headerRow) {
                    const headers = Object.values(headerRow);
                    const headerIndex = excelData.indexOf(headerRow);
                    excelData = excelData.slice(headerIndex + 1).map(row => {
                        const obj = {};
                        headers.forEach((header, i) => {
                            obj[header.trim()] = row[`__EMPTY_${i+1}`] || '';
                        });
                        return obj;
                    });
                }
                console.log('Excel data loaded:', excelData);
                updateGenerateButton();
            } catch (error) {
                console.error('Error reading Excel file:', error);
                alert('Error reading Excel file. Please make sure it is a valid Excel file.');
            }
        };
        reader.onerror = function() {
            alert('Error reading the file. Please try again.');
        };
        reader.readAsArrayBuffer(file);
    }
}

function handleImageUpload(event) {
    const files = event.target.files;
    images.clear();
    if (files.length > 0) {
        document.getElementById('imageInfo').textContent = `Loaded ${files.length} images`;
        Array.from(files).forEach((file) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                const id = file.name.replace(/\.[^/.]+$/, "").toLowerCase();
                if (id === 'f') {
                    defaultFemaleImage = e.target.result;
                } else if (id === 'm') {
                    defaultMaleImage = e.target.result;
                } else {
                    images.set(id, e.target.result);
                }
                updateGenerateButton();
                if (typeof renderUploadedImagesList === 'function') renderUploadedImagesList();
            };
            reader.onerror = function() {
                console.error('Error loading image:', file.name);
            };
            reader.readAsDataURL(file);
        });
    } else {
        document.getElementById('imageInfo').textContent = '';
        updateGenerateButton();
        if (typeof renderUploadedImagesList === 'function') renderUploadedImagesList();
    }
}

function handleLogoUpload(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('logoInfo').textContent = file.name;
        const reader = new FileReader();
        reader.onload = function(e) {
            uploadedLogoBase64 = e.target.result;
        };
        reader.readAsDataURL(file);
    }
}

function updateGenerateButton() {
    const generateBtn = document.getElementById('generateBtn');
    const hasExcelData = excelData && excelData.length > 0;
    const hasImages = images.size > 0;
    
    console.log('Update button state:', { hasExcelData, hasImages });
    generateBtn.disabled = !(hasExcelData && hasImages);
    
    if (!generateBtn.disabled) {
        generateBtn.style.backgroundColor = '#4CAF50';
    } else {
        generateBtn.style.backgroundColor = '#cccccc';
    }
}

// Helper to add leading zero if not present
function withLeadingZero(num) {
    if (!num) return '';
    num = String(num).trim();
    return num.startsWith('0') ? num : '0' + num;
}

function createIDCard(student, index) {
    const card = document.createElement('div');
    card.className = 'id-card';

    // Header
    const header = document.createElement('div');
    header.className = 'id-card-header';

    // Logo
    const logo = document.createElement('img');
    logo.className = 'id-card-logo';
    logo.src = uploadedLogoBase64 || LOGO_BASE64;
    logo.alt = 'Logo';
    header.appendChild(logo);

    // ID badge
    const idBadge = document.createElement('div');
    idBadge.className = 'id-card-id';
    idBadge.textContent = String(student['ADHERENT'] || '');
    header.appendChild(idBadge);

    card.appendChild(header);

    // Photo wrapper (flexbox)
    const photoWrapper = document.createElement('div');
    photoWrapper.className = 'id-card-photo-wrapper';
    // Student photo by ID or fallback to type
    const id = String(student['ADHERENT'] || '').trim().toLowerCase();
    let photoSrc = images.get(id) || '';
    if (!photoSrc) {
        let type = (student['type'] || student['TYPE'] || '').toString().trim().toLowerCase();
        if (type === 'f' && defaultFemaleImage) {
            photoSrc = defaultFemaleImage;
        } else if (type === 'm' && defaultMaleImage) {
            photoSrc = defaultMaleImage;
        }
    }
    if (photoSrc) {
        const photo = document.createElement('img');
        photo.className = 'id-card-photo';
        photo.src = photoSrc;
        photo.alt = 'Student Photo';
        photoWrapper.appendChild(photo);
    }
    card.appendChild(photoWrapper);

    // Card content
    const content = document.createElement('div');
    content.className = 'id-card-content';

    // Full Name with numbering (with leading zero)
    const cardNumber = (index + 1).toString().padStart(2, '0');
    const name = document.createElement('div');
    name.className = 'id-card-name';
    name.textContent = `${cardNumber}. ${String(student['NOM & PRENOM'] || '')}`;
    content.appendChild(name);

    // TEL ADH
    const telAdh = document.createElement('div');
    telAdh.className = 'id-card-info-row';
    telAdh.innerHTML = `
        <span class="id-card-label">TEL ADH :</span>
        <span class="id-card-value">${withLeadingZero(student['TEL ADH'])}</span>
    `;
    content.appendChild(telAdh);

    // TEL PARENT
    const telParent = document.createElement('div');
    telParent.className = 'id-card-info-row';
    telParent.innerHTML = `
        <span class="id-card-label">TEL PARENT :</span>
        <span class="id-card-value">${withLeadingZero(student['TEL PARENT'])}</span>
    `;
    content.appendChild(telParent);

    // LYCEE
    const lycee = document.createElement('div');
    lycee.className = 'id-card-info-row';
    lycee.innerHTML = `
        <span class="id-card-label">LYCEE :</span>
        <span class="id-card-value">${String(student['LYCEE'] || '')}</span>
    `;
    content.appendChild(lycee);

    card.appendChild(content);

    return card;
}

async function generatePDF() {
    if (!excelData || excelData.length === 0) {
        alert('No Excel data loaded. Please upload an Excel file.');
        return;
    }
    if (images.size === 0) {
        alert('No images loaded. Please upload student photos first.');
        return;
    }

    console.log('Starting PDF generation...');
    console.log('Excel data:', excelData);
    console.log('Images loaded:', images.size);

    // Show loading indicator
    const loading = document.createElement('div');
    loading.textContent = 'Generating PDF, please wait...';
    loading.style.textAlign = 'center';
    loading.style.padding = '20px';
    loading.style.fontSize = '16px';
    loading.style.color = '#1a73e8';
    document.body.appendChild(loading);

    // Filter out empty rows
    const filteredData = excelData.filter(student =>
        student['NOM & PRENOM'] ||
        student['TEL ADH'] ||
        student['TEL PARENT'] ||
        student['LYCEE'] ||
        student['ADHERENT']
    );

    console.log('Filtered data:', filteredData.length, 'cards');

    // Show total cards
    document.getElementById('cardCount').textContent = `Total cards generated: ${filteredData.length}`;

    // Create preview
    const preview = document.getElementById('preview');
    preview.innerHTML = '';
    preview.style.overflowX = 'auto';
    preview.style.overflowY = 'auto';
    preview.style.maxHeight = '80vh';
    preview.style.width = '100%';
    preview.style.maxWidth = '100%';

    const container = document.createElement('div');
    container.style.display = 'grid';
    container.style.gridTemplateColumns = 'repeat(2, 1fr)';
    container.style.gap = '10px';
    container.style.padding = '10px';
    preview.appendChild(container);

    // Create all cards first
    const cards = filteredData.map((student, index) => {
        const card = createIDCard(student, index);
        container.appendChild(card);
        return card;
    });

    console.log('Created', cards.length, 'preview cards');

    // Wait for all images to load
    const imagePromises = Array.from(container.getElementsByTagName('img')).map(img => {
        return new Promise((resolve) => {
            if (img.complete) {
                resolve();
            } else {
                img.onload = resolve;
                img.onerror = () => {
                    console.error('Error loading image:', img.src);
                    resolve();
                };
            }
        });
    });

    try {
        // Wait for all images to load
        await Promise.all(imagePromises);
        console.log('All preview images loaded');

        // Create PDF
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('portrait', 'mm', 'a4');
        const cardsPerPage = 8;
        const cardWidth = 97; // mm
        const cardHeight = 65; // mm
        const marginX = 2; // mm
        const marginY = 4; // mm
        const pageMarginX = 7; // mm
        const pageMarginY = 7; // mm

        let currentRow = 0;
        let currentCol = 0;
        let pageCount = 0;

        // Render each card to canvas and add to PDF
        for (let i = 0; i < cards.length; i++) {
            // Add new page if needed
            if (i > 0 && i % cardsPerPage === 0) {
                pdf.addPage();
                pageCount++;
                currentRow = 0;
                currentCol = 0;
            }
            const x = pageMarginX + (currentCol * (cardWidth + marginX));
            const y = pageMarginY + (currentRow * (cardHeight + marginY));

            // Render card to canvas
            const canvas = await html2canvas(cards[i], {
                scale: 3,
                useCORS: true,
                backgroundColor: '#fff',
                logging: false
            });
            const imgData = canvas.toDataURL('image/jpeg', 1.0);
            pdf.addImage(imgData, 'JPEG', x, y, cardWidth, cardHeight);

            currentRow++;
            if (currentRow >= 4) {
                currentRow = 0;
                currentCol++;
            }
        }

        // Remove loading indicator
        document.body.removeChild(loading);

        // Save the PDF
        pdf.save('id_cards.pdf');
        console.log('PDF saved successfully');
    } catch (error) {
        console.error('Error generating PDF:', error);
        alert('An error occurred while generating the PDF. Please check console for details.');
        document.body.removeChild(loading);
    }
}