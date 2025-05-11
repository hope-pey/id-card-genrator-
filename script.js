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

function withLeadingZero(num) {
    if (!num) return '';
    num = String(num).trim();
    return num.startsWith('0') ? num : '0' + num;
}

function createIDCard(student, index) {
    const card = document.createElement('div');
    card.className = 'id-card';

 
    const header = document.createElement('div');
    header.className = 'id-card-header';

    const logo = document.createElement('img');
    logo.className = 'id-card-logo';
    logo.src = uploadedLogoBase64 || LOGO_BASE64;
    logo.alt = 'Logo';
    header.appendChild(logo);

    const idBadge = document.createElement('div');
    idBadge.className = 'id-card-id';
    idBadge.textContent = String(student['ADHERENT'] || '');
    header.appendChild(idBadge);

    card.appendChild(header);
    const photoWrapper = document.createElement('div');
    photoWrapper.className = 'id-card-photo-wrapper';
    

    const id = String(student['ADHERENT'] || '').trim().toLowerCase();
    const type = String(student['type'] || student['TYPE'] || '').trim().toLowerCase();
    

    let photoSrc = images.get(id);
    if (!photoSrc) {
        if (type === 'f') {
            photoSrc = defaultFemaleImage;
        } else if (type === 'm') {
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

    const content = document.createElement('div');
    content.className = 'id-card-content';

    const cardNumber = (index + 1).toString().padStart(2, '0');
    const name = document.createElement('div');
    name.className = 'id-card-name';
    name.textContent = `${cardNumber}. ${String(student['NOM & PRENOM'] || '')}`;
    content.appendChild(name);

    const telAdh = document.createElement('div');
    telAdh.className = 'id-card-info-row';
    telAdh.innerHTML = `
        <span class="id-card-label">TEL ADH :</span>
        <span class="id-card-value">${withLeadingZero(student['TEL ADH'])}</span>
    `;
    content.appendChild(telAdh);

    const telParent = document.createElement('div');
    telParent.className = 'id-card-info-row';
    telParent.innerHTML = `
        <span class="id-card-label">TEL PARENT :</span>
        <span class="id-card-value">${withLeadingZero(student['TEL PARENT'])}</span>
    `;
    content.appendChild(telParent);


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
    const modalOverlay = document.createElement('div');
    modalOverlay.id = 'pdf-modal-overlay';
    modalOverlay.style.position = 'fixed';
    modalOverlay.style.top = '0';
    modalOverlay.style.left = '0';
    modalOverlay.style.width = '100%';
    modalOverlay.style.height = '100%';
    modalOverlay.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    modalOverlay.style.display = 'flex';
    modalOverlay.style.justifyContent = 'center';
    modalOverlay.style.alignItems = 'center';
    modalOverlay.style.zIndex = '9999';

    const modalContent = document.createElement('div');
    modalContent.id = 'pdf-modal';
    modalContent.style.backgroundColor = 'white';
    modalContent.style.padding = '2rem';
    modalContent.style.borderRadius = '8px';
    modalContent.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.1)';
    modalContent.style.textAlign = 'center';
    modalContent.style.maxWidth = '90%';
    modalContent.style.width = 'auto';

    const loadingText = document.createElement('div');
    loadingText.textContent = 'Generating PDF, please wait...';
    loadingText.style.fontSize = '1.2rem';
    loadingText.style.color = '#1a73e8';
    loadingText.style.marginBottom = '1rem';

    const spinner = document.createElement('div');
    spinner.style.border = '4px solid #f3f3f3';
    spinner.style.borderTop = '4px solid #1a73e8';
    spinner.style.borderRadius = '50%';
    spinner.style.width = '40px';
    spinner.style.height = '40px';
    spinner.style.animation = 'spin 1s linear infinite';
    spinner.style.margin = '0 auto';


    const style = document.createElement('style');
    style.textContent = `
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    `;
    document.head.appendChild(style);

    modalContent.appendChild(loadingText);
    modalContent.appendChild(spinner);
    modalOverlay.appendChild(modalContent);
    document.body.appendChild(modalOverlay);

    const filteredData = excelData.filter(student =>
        student['NOM & PRENOM'] ||
        student['TEL ADH'] ||
        student['TEL PARENT'] ||
        student['LYCEE'] ||
        student['ADHERENT']
    );

    console.log('Filtered data:', filteredData.length, 'cards');

    document.getElementById('cardCount').textContent = `Total cards generated: ${filteredData.length}`;

 
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

    const cards = filteredData.map((student, index) => {
        const card = createIDCard(student, index);
        container.appendChild(card);
        return card;
    });

    console.log('Created', cards.length, 'preview cards');


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

        await Promise.all(imagePromises);
        console.log('All preview images loaded');

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('portrait', 'mm', 'a4');
        const cardsPerPage = 8;
        const cardWidth = 97; 
        const cardHeight = 65; 
        const marginX = 2; 
        const marginY = 4; 
        const pageMarginX = 7; 
        const pageMarginY = 7; 

        let currentRow = 0;
        let currentCol = 0;
        let pageCount = 0;

        for (let i = 0; i < cards.length; i++) {
            if (i > 0 && i % cardsPerPage === 0) {
                pdf.addPage();
                pageCount++;
                currentRow = 0;
                currentCol = 0;
            }
            const x = pageMarginX + (currentCol * (cardWidth + marginX));
            const y = pageMarginY + (currentRow * (cardHeight + marginY));
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
        document.body.removeChild(modalOverlay);
        document.head.removeChild(style);
        pdf.save('id_cards.pdf');
        console.log('PDF saved successfully');
    } catch (error) {
        console.error('Error generating PDF:', error);
        alert('An error occurred while generating the PDF. Please check console for details.');
        document.body.removeChild(modalOverlay);
        document.head.removeChild(style);
    }
} 