// Local Storage Keys
const QUOTATIONS_KEY = 'quotations_data';
const INVOICES_KEY = 'invoices_data';
const FORM_DRAFT_KEY = 'document_form_draft';

// Global variables
let itemCounter = 1;
let currentDocumentData = null;
let allQuotations = [];
let allInvoices = [];
let currentDocumentType = 'quotation';

// DOM Elements
const form = document.getElementById('documentForm');
const statusIndicator = document.getElementById('statusIndicator');
const notification = document.getElementById('notification');
const clearBtn = document.getElementById('clearBtn');
const addItemBtn = document.getElementById('addItemBtn');
const itemsContainer = document.getElementById('itemsContainer');
const previewSection = document.getElementById('previewSection');
const documentPreview = document.getElementById('documentPreview');
const printBtn = document.getElementById('printBtn');
const editBtn = document.getElementById('editBtn');
const generateBtn = document.getElementById('generateBtn');
const previewBtn = document.getElementById('previewBtn');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const mobileMenuToggle = document.getElementById('mobileMenuToggle');
const sidebar = document.getElementById('sidebar');

// Tab Elements
const tabBtns = document.querySelectorAll('.tab-btn');
const mainTitle = document.getElementById('mainTitle');
const mainSubtitle = document.getElementById('mainSubtitle');
const formTitle = document.getElementById('formTitle');
const previewTitle = document.getElementById('previewTitle');

// Sidebar Elements
const saveDataBtn = document.getElementById('saveDataBtn');
const exportAllBtn = document.getElementById('exportAllBtn');
const clearAllBtn = document.getElementById('clearAllBtn');
const newDocumentBtn = document.getElementById('newDocumentBtn');
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');
const recentList = document.getElementById('recentList');
const totalDocuments = document.getElementById('totalDocuments');
const totalAmount = document.getElementById('totalAmount');

// Payment Elements
const paymentSection = document.getElementById('paymentSection');
const paymentFields = document.getElementById('paymentFields');
const validTillGroup = document.getElementById('validTillGroup');
const dueDateGroup = document.getElementById('dueDateGroup');
const advanceAmount = document.getElementById('advanceAmount');
const advanceRow = document.getElementById('advanceRow');
const advanceDisplay = document.getElementById('advanceDisplay');

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    loadStoredData();
    generateDocumentNumber();
    setDefaultDates();
    calculateTotals();
    setStatus('ready', 'Ready');
    updateSidebarStats();
    updateRecentList();
    updateUIForDocumentType();
});

/**
 * Initialize all event listeners
 */
function initializeEventListeners() {
    // Tab switching
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => switchDocumentType(btn.dataset.type));
    });

    // Form actions
    generateBtn.addEventListener('click', generateDocument);
    previewBtn.addEventListener('click', showPreview);
    downloadExcelBtn.addEventListener('click', downloadCurrentExcel);
    clearBtn.addEventListener('click', clearForm);
    addItemBtn.addEventListener('click', addNewItem);
    
    // Preview actions
    printBtn.addEventListener('click', printDocument);
    editBtn.addEventListener('click', editDocument);
    
    // Sidebar actions
    saveDataBtn.addEventListener('click', saveCurrentDocument);
    exportAllBtn.addEventListener('click', exportAllToExcel);
    clearAllBtn.addEventListener('click', clearAllData);
    newDocumentBtn.addEventListener('click', newDocument);
    
    // Mobile menu toggle
    mobileMenuToggle.addEventListener('click', toggleMobileMenu);
    
    // Item calculations
    itemsContainer.addEventListener('input', handleItemCalculation);
    itemsContainer.addEventListener('click', handleItemRemoval);
    
    // Search functionality
    searchInput.addEventListener('input', handleSearch);
    
    // Payment calculations
    if (advanceAmount) {
        advanceAmount.addEventListener('input', calculateTotals);
    }
    
    // Auto-save draft functionality
    form.addEventListener('input', debounce(saveDraftToLocal, 1000));
    
    // Close sidebar when clicking outside on mobile
    document.addEventListener('click', function(e) {
        if (window.innerWidth <= 992 && sidebar.classList.contains('active')) {
            if (!sidebar.contains(e.target) && !mobileMenuToggle.contains(e.target)) {
                sidebar.classList.remove('active');
            }
        }
    });
    
    // Handle window resize
    window.addEventListener('resize', handleWindowResize);
}

/**
 * Switch between document types
 */
function switchDocumentType(type) {
    if (currentDocumentType === type) return;
    
    // Confirm switch if form has data
    if (hasFormData() && !confirm('Switching document type will clear current form data. Continue?')) {
        return;
    }
    
    currentDocumentType = type;
    
    // Update tab buttons
    tabBtns.forEach(btn => btn.classList.remove('active'));
    document.querySelector(`[data-type="${type}"]`).classList.add('active');
    
    // Clear form and reset
    clearForm(false);
    updateUIForDocumentType();
    updateSidebarStats();
    updateRecentList();
    
    showNotification(`Switched to ${type} mode`, 'info');
}

/**
 * Update UI based on document type
 */
function updateUIForDocumentType() {
    const isInvoice = currentDocumentType === 'invoice';
    
    // Update titles and labels
    mainTitle.textContent = isInvoice ? 'Professional Invoice System' : 'Professional Quotation System';
    mainSubtitle.textContent = isInvoice ? 'Generate & Store Invoices with Excel Export' : 'Generate & Store Quotations with Excel Export';
    formTitle.textContent = isInvoice ? 'Create New Invoice' : 'Create New Quotation';
    previewTitle.textContent = isInvoice ? 'Invoice Preview' : 'Quotation Preview';
    
    document.getElementById('documentNoLabel').textContent = isInvoice ? 'Invoice No.' : 'Quotation No.';
    document.getElementById('documentDateLabel').textContent = isInvoice ? 'Invoice Date' : 'Quotation Date';
    document.getElementById('itemsSectionTitle').textContent = isInvoice ? 'Invoice Items' : 'Quotation Items';
    document.getElementById('generateBtnText').textContent = isInvoice ? 'Generate Invoice' : 'Generate Quotation';
    document.getElementById('newDocumentText').textContent = isInvoice ? 'New Invoice' : 'New Quotation';
    document.getElementById('documentsLabel').textContent = isInvoice ? 'Total Invoices' : 'Total Quotations';
    document.getElementById('recentDocumentsTitle').textContent = isInvoice ? 'Recent Invoices' : 'Recent Quotations';
    
    // Show/hide payment fields for invoices
    if (paymentSection && paymentFields) {
        paymentSection.style.display = isInvoice ? 'flex' : 'none';
        paymentFields.style.display = isInvoice ? 'grid' : 'none';
    }
    
    // Show/hide date fields
    if (validTillGroup && dueDateGroup) {
        validTillGroup.style.display = isInvoice ? 'none' : 'flex';
        dueDateGroup.style.display = isInvoice ? 'flex' : 'none';
    }
    
    // Update grand total label
    document.getElementById('grandTotalLabel').textContent = isInvoice ? 'Amount Due:' : 'Grand Total:';
    
    generateDocumentNumber();
    setDefaultDates();
}

/**
 * Check if form has data
 */
function hasFormData() {
    const companyName = document.getElementById('companyName').value.trim();
    const clientName = document.getElementById('clientName').value.trim();
    return companyName || clientName;
}

/**
 * Toggle mobile menu
 */
function toggleMobileMenu() {
    sidebar.classList.toggle('active');
}

/**
 * Handle window resize
 */
function handleWindowResize() {
    if (window.innerWidth > 992) {
        sidebar.classList.remove('active');
    }
}

/**
 * Load stored data from localStorage
 */
function loadStoredData() {
    try {
        const quotationsData = localStorage.getItem(QUOTATIONS_KEY);
        const invoicesData = localStorage.getItem(INVOICES_KEY);
        
        if (quotationsData) {
            allQuotations = JSON.parse(quotationsData);
        }
        
        if (invoicesData) {
            allInvoices = JSON.parse(invoicesData);
        }
        
        console.log(`Loaded ${allQuotations.length} quotations and ${allInvoices.length} invoices from storage`);
        
        // Load form draft
        const draftData = localStorage.getItem(FORM_DRAFT_KEY);
        if (draftData) {
            const draft = JSON.parse(draftData);
            if (draft.documentType === currentDocumentType) {
                loadFormDraft(draft);
            }
        }
    } catch (error) {
        console.error('Error loading stored data:', error);
        allQuotations = [];
        allInvoices = [];
    }
}

/**
 * Generate unique document number
 */
function generateDocumentNumber() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const time = String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0');
    
    const prefix = currentDocumentType === 'invoice' ? 'INV' : 'QUO';
    const documentNo = `${prefix}${year}${month}${day}${time}`;
    document.getElementById('documentNo').value = documentNo;
}

/**
 * Set default dates
 */
function setDefaultDates() {
    const today = new Date();
    const futureDate = new Date();
    
    if (currentDocumentType === 'quotation') {
        futureDate.setDate(today.getDate() + 30); // 30 days validity
        document.getElementById('validTillDate').value = futureDate.toISOString().split('T')[0];
    } else {
        futureDate.setDate(today.getDate() + 30); // 30 days due date
        document.getElementById('dueDate').value = futureDate.toISOString().split('T')[0];
    }
    
    document.getElementById('documentDate').value = today.toISOString().split('T')[0];
}

/**
 * Add new item row
 */
function addNewItem() {
    itemCounter++;
    const itemRow = document.createElement('div');
    itemRow.className = 'item-row hover-lift';
    itemRow.setAttribute('data-item', itemCounter);
    
    itemRow.innerHTML = `
        <div class="item-grid">
            <div class="form-group item-description">
                <label>Item Description</label>
                <input type="text" name="itemDescription[]" placeholder="Enter item description" required>
            </div>
            <div class="form-group">
                <label>GST Rate (%)</label>
                <select name="gstRate[]" class="gst-rate">
                    <option value="0">0%</option>
                    <option value="5">5%</option>
                    <option value="12">12%</option>
                    <option value="18" selected>18%</option>
                    <option value="28">28%</option>
                </select>
            </div>
            <div class="form-group">
                <label>Quantity</label>
                <input type="number" name="quantity[]" min="1" value="1" class="quantity" required>
            </div>
            <div class="form-group">
                <label>Rate (₹)</label>
                <input type="number" name="rate[]" min="0" step="0.01" class="rate" placeholder="0.00" required>
            </div>
            <div class="form-group">
                <label>Amount (₹)</label>
                <input type="number" name="amount[]" class="amount" readonly>
            </div>
            <div class="form-group">
                <label>CGST (₹)</label>
                <input type="number" name="cgst[]" class="cgst" readonly>
            </div>
            <div class="form-group">
                <label>SGST (₹)</label>
                <input type="number" name="sgst[]" class="sgst" readonly>
            </div>
            <div class="form-group">
                <label>Total (₹)</label>
                <input type="number" name="total[]" class="item-total" readonly>
            </div>
            <div class="form-group item-actions">
                <label>&nbsp;</label>
                <button type="button" class="btn btn-danger btn-sm remove-item">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        </div>
    `;
    
    itemsContainer.appendChild(itemRow);
    showNotification('New item added successfully', 'success');
}

/**
 * Handle item calculations
 */
function handleItemCalculation(e) {
    const itemRow = e.target.closest('.item-row');
    if (!itemRow) return;
    
    const quantity = parseFloat(itemRow.querySelector('.quantity').value) || 0;
    const rate = parseFloat(itemRow.querySelector('.rate').value) || 0;
    const gstRate = parseFloat(itemRow.querySelector('.gst-rate').value) || 0;
    
    const amount = quantity * rate;
    const cgst = (amount * gstRate) / 200; // CGST is half of GST
    const sgst = (amount * gstRate) / 200; // SGST is half of GST
    const total = amount + cgst + sgst;
    
    itemRow.querySelector('.amount').value = amount.toFixed(2);
    itemRow.querySelector('.cgst').value = cgst.toFixed(2);
    itemRow.querySelector('.sgst').value = sgst.toFixed(2);
    itemRow.querySelector('.item-total').value = total.toFixed(2);
    
    calculateTotals();
}

/**
 * Handle item removal
 */
function handleItemRemoval(e) {
    if (e.target.closest('.remove-item')) {
        const itemRow = e.target.closest('.item-row');
        const itemRows = itemsContainer.querySelectorAll('.item-row');
        
        if (itemRows.length > 1) {
            itemRow.remove();
            calculateTotals();
            showNotification('Item removed successfully', 'info');
        } else {
            showNotification('At least one item is required', 'warning');
        }
    }
}

/**
 * Calculate totals
 */
function calculateTotals() {
    const amounts = Array.from(document.querySelectorAll('.amount')).map(input => parseFloat(input.value) || 0);
    const cgsts = Array.from(document.querySelectorAll('.cgst')).map(input => parseFloat(input.value) || 0);
    const sgsts = Array.from(document.querySelectorAll('.sgst')).map(input => parseFloat(input.value) || 0);
    
    const subtotal = amounts.reduce((sum, amount) => sum + amount, 0);
    const totalCGST = cgsts.reduce((sum, cgst) => sum + cgst, 0);
    const totalSGST = sgsts.reduce((sum, sgst) => sum + sgst, 0);
    const grandTotal = subtotal + totalCGST + totalSGST;
    
    // Handle advance payment for invoices
    let advance = 0;
    let amountDue = grandTotal;
    
    if (currentDocumentType === 'invoice' && advanceAmount) {
        advance = parseFloat(advanceAmount.value) || 0;
        amountDue = grandTotal - advance;
        
        if (advanceRow && advanceDisplay) {
            advanceRow.style.display = advance > 0 ? 'flex' : 'none';
            advanceDisplay.textContent = `₹${advance.toFixed(2)}`;
        }
    }
    
    document.getElementById('subtotal').textContent = `₹${subtotal.toFixed(2)}`;
    document.getElementById('totalCGST').textContent = `₹${totalCGST.toFixed(2)}`;
    document.getElementById('totalSGST').textContent = `₹${totalSGST.toFixed(2)}`;
    document.getElementById('grandTotal').textContent = `₹${(currentDocumentType === 'invoice' ? amountDue : grandTotal).toFixed(2)}`;
}

/**
 * Generate document
 */
function generateDocument() {
    if (!validateForm()) {
        return;
    }
    
    const formData = getFormData();
    
    // Check for duplicates
    const isDuplicate = checkForDuplicates(formData);
    if (isDuplicate) {
        setStatus('duplicate', 'Duplicate found');
        showNotification(`A ${currentDocumentType} with this client email already exists. Please check existing records.`, 'warning');
        return;
    }
    
    currentDocumentData = formData;
    generateDocumentPreview(formData);
    showPreview();
    setStatus('ready', `${currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1)} generated`);
    showNotification(`${currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1)} generated successfully! Use "Save to Storage" to save it.`, 'success');
}

/**
 * Save current document to storage
 */
function saveCurrentDocument() {
    if (!currentDocumentData) {
        showNotification(`Please generate a ${currentDocumentType} first`, 'warning');
        return;
    }
    
    try {
        setStatus('saving', 'Saving...');
        
        // Add to appropriate array
        if (currentDocumentType === 'quotation') {
            allQuotations.push(currentDocumentData);
            localStorage.setItem(QUOTATIONS_KEY, JSON.stringify(allQuotations));
        } else {
            allInvoices.push(currentDocumentData);
            localStorage.setItem(INVOICES_KEY, JSON.stringify(allInvoices));
        }
        
        // Clear form draft
        localStorage.removeItem(FORM_DRAFT_KEY);
        
        updateSidebarStats();
        updateRecentList();
        setStatus('saved', 'Saved successfully');
        
        const totalCount = currentDocumentType === 'quotation' ? allQuotations.length : allInvoices.length;
        showNotification(`${currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1)} saved! Total: ${totalCount} ${currentDocumentType}s stored`, 'success');
        
        return true;
    } catch (error) {
        console.error('Error saving to localStorage:', error);
        setStatus('error', 'Save failed');
        showNotification(`Error saving ${currentDocumentType} data`, 'error');
        return false;
    }
}

/**
 * Download current document as Excel
 */
function downloadCurrentExcel() {
    if (!currentDocumentData) {
        showNotification(`Please generate a ${currentDocumentType} first`, 'warning');
        return;
    }
    
    try {
        setStatus('saving', 'Creating Excel file...');
        
        const wb = XLSX.utils.book_new();
        const documentSheet = createDocumentSheet(currentDocumentData);
        XLSX.utils.book_append_sheet(wb, documentSheet, currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1));
        
        const fileName = `${currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1)}_${currentDocumentData.documentNo}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        setStatus('ready', 'Ready');
        showNotification('Excel file downloaded successfully!', 'success');
    } catch (error) {
        console.error('Error creating Excel file:', error);
        setStatus('error', 'Export failed');
        showNotification('Error creating Excel file', 'error');
    }
}

/**
 * Export all documents to Excel
 */
function exportAllToExcel() {
    const currentData = currentDocumentType === 'quotation' ? allQuotations : allInvoices;
    
    if (currentData.length === 0) {
        showNotification(`No ${currentDocumentType}s to export`, 'warning');
        return;
    }
    
    try {
        setStatus('saving', 'Exporting all data...');
        
        const wb = XLSX.utils.book_new();
        const docTypeName = currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1);
        
        // Create summary sheet
        const summaryData = [
            [`${docTypeName.toUpperCase()} SUMMARY REPORT`],
            ['Generated on:', new Date().toLocaleString('en-IN')],
            [`Total ${docTypeName}s:`, currentData.length],
            ['Total Amount:', `₹${currentData.reduce((sum, q) => sum + (q.grandTotal || 0), 0).toFixed(2)}`],
            [],
            [`${docTypeName} No`, 'Date', 'Company', 'Client Name', 'Client Email', 'Items Count', 'Subtotal (₹)', 'CGST (₹)', 'SGST (₹)', 'Grand Total (₹)']
        ];
        
        if (currentDocumentType === 'invoice') {
            summaryData[5].push('Payment Status', 'Amount Due (₹)');
        }
        
        currentData.forEach(doc => {
            const row = [
                doc.documentNo,
                doc.documentDate,
                doc.companyName,
                doc.clientName,
                doc.clientEmail,
                doc.items.length,
                doc.subtotal.toFixed(2),
                doc.totalCGST.toFixed(2),
                doc.totalSGST.toFixed(2),
                doc.grandTotal.toFixed(2)
            ];
            
            if (currentDocumentType === 'invoice') {
                row.push(doc.paymentStatus || 'unpaid');
                row.push((doc.grandTotal - (doc.advanceAmount || 0)).toFixed(2));
            }
            
            summaryData.push(row);
        });
        
        const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
        const maxWidth = 25;
        const wscols = summaryData[0].map(() => ({ wch: maxWidth }));
        summarySheet['!cols'] = wscols;
        
        XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');
        
        // Create individual sheets (limit to first 15 for performance)
        const maxSheets = Math.min(currentData.length, 15);
        for (let i = 0; i < maxSheets; i++) {
            const doc = currentData[i];
            const docSheet = createDocumentSheet(doc);
            const sheetName = `${docTypeName.charAt(0)}${i + 1}_${doc.documentNo.slice(-8)}`;
            XLSX.utils.book_append_sheet(wb, docSheet, sheetName);
        }
        
        const fileName = `All_${docTypeName}s_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        setStatus('ready', 'Ready');
        showNotification(`Exported ${currentData.length} ${currentDocumentType}s to Excel successfully!`, 'success');
        
    } catch (error) {
        console.error('Error exporting to Excel:', error);
        setStatus('error', 'Export failed');
        showNotification('Error exporting data to Excel', 'error');
    }
}

/**
 * Create Excel sheet for a document
 */
function createDocumentSheet(data) {
    const docTypeName = (data.documentType || currentDocumentType).charAt(0).toUpperCase() + (data.documentType || currentDocumentType).slice(1);
    const sheetData = [];
    
    // Header information
    sheetData.push([data.companyName.toUpperCase()]);
    sheetData.push([docTypeName.toUpperCase()]);
    sheetData.push([]);
    sheetData.push(['Company Information:']);
    sheetData.push(['Name:', data.companyName]);
    sheetData.push(['Address:', data.companyAddress.replace(/\n/g, ', ')]);
    sheetData.push([]);
    sheetData.push([`${docTypeName} Details:`]);
    sheetData.push([`${docTypeName} No:`, data.documentNo]);
    sheetData.push([`${docTypeName} Date:`, formatDateForExcel(data.documentDate)]);
    
    if (data.documentType === 'quotation' || currentDocumentType === 'quotation') {
        sheetData.push(['Valid Till:', formatDateForExcel(data.validTillDate)]);
    } else {
        sheetData.push(['Due Date:', formatDateForExcel(data.dueDate || data.validTillDate)]);
        if (data.paymentStatus) {
            sheetData.push(['Payment Status:', data.paymentStatus.charAt(0).toUpperCase() + data.paymentStatus.slice(1)]);
        }
        if (data.paymentTerms) {
            sheetData.push(['Payment Terms:', data.paymentTerms]);
        }
    }
    
    sheetData.push([]);
    sheetData.push(['Client Information:']);
    sheetData.push(['Name:', data.clientName]);
    sheetData.push(['Email:', data.clientEmail]);
    sheetData.push(['Phone:', data.clientPhone || 'Not provided']);
    sheetData.push(['Location:', data.clientLocation || 'Not provided']);
    sheetData.push([]);
    
    // Items header
    sheetData.push(['ITEMS DETAILS']);
    sheetData.push(['S.No', 'Description', 'GST Rate (%)', 'Quantity', 'Rate (₹)', 'Amount (₹)', 'CGST (₹)', 'SGST (₹)', 'Total (₹)']);
    
    // Items data
    data.items.forEach((item, index) => {
        sheetData.push([
            index + 1,
            item.description,
            `${item.gstRate}%`,
            item.quantity,
            item.rate.toFixed(2),
            item.amount.toFixed(2),
            item.cgst.toFixed(2),
            item.sgst.toFixed(2),
            item.total.toFixed(2)
        ]);
    });
    
    // Totals
    sheetData.push([]);
    sheetData.push(['', '', '', '', '', 'Subtotal:', data.subtotal.toFixed(2)]);
    sheetData.push(['', '', '', '', '', 'Total CGST:', data.totalCGST.toFixed(2)]);
    sheetData.push(['', '', '', '', '', 'Total SGST:', data.totalSGST.toFixed(2)]);
    
    if (data.advanceAmount && data.advanceAmount > 0) {
        sheetData.push(['', '', '', '', '', 'Advance Received:', data.advanceAmount.toFixed(2)]);
        sheetData.push(['', '', '', '', '', 'AMOUNT DUE:', (data.grandTotal - data.advanceAmount).toFixed(2)]);
    } else {
        sheetData.push(['', '', '', '', '', 'GRAND TOTAL:', data.grandTotal.toFixed(2)]);
    }
    
    sheetData.push([]);
    sheetData.push(['Terms & Conditions:']);
    
    const termsLines = data.termsConditions.split('\n');
    termsLines.forEach(line => {
        if (line.trim()) {
            sheetData.push([line.trim()]);
        }
    });
    
    sheetData.push([]);
    sheetData.push(['Generated on:', new Date().toLocaleString('en-IN')]);
    
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Set column widths
    worksheet['!cols'] = [
        { wch: 8 },  // S.No
        { wch: 35 }, // Description
        { wch: 12 }, // GST Rate
        { wch: 10 }, // Quantity
        { wch: 12 }, // Rate
        { wch: 12 }, // Amount
        { wch: 12 }, // CGST
        { wch: 12 }, // SGST
        { wch: 12 }  // Total
    ];
    
    return worksheet;
}

/**
 * Format date for Excel display
 */
function formatDateForExcel(dateString) {
    if (!dateString) return 'Not specified';
    const date = new Date(dateString);
    return date.toLocaleDateString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
    });
}

/**
 * Validate form data
 */
function validateForm() {
    const requiredFields = [
        'documentDate', 'companyName', 'companyAddress',
        'clientName', 'clientEmail'
    ];
    
    // Add document type specific required fields
    if (currentDocumentType === 'quotation') {
        requiredFields.push('validTillDate');
    } else {
        requiredFields.push('dueDate');
    }
    
    let isValid = true;
    let firstInvalidField = null;
    
    // Check required fields
    requiredFields.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (!field) return;
        
        const value = field.value.trim();
        
        if (!value) {
            field.style.borderColor = '#dc2626';
            field.style.boxShadow = '0 0 0 3px rgba(220, 38, 38, 0.1)';
            if (!firstInvalidField) firstInvalidField = field;
            isValid = false;
        } else {
            field.style.borderColor = '#e5e7eb';
            field.style.boxShadow = 'none';
        }
    });
    
    // Validate email format
    const email = document.getElementById('clientEmail').value;
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (email && !emailRegex.test(email)) {
        const emailField = document.getElementById('clientEmail');
        emailField.style.borderColor = '#dc2626';
        emailField.style.boxShadow = '0 0 0 3px rgba(220, 38, 38, 0.1)';
        if (!firstInvalidField) firstInvalidField = emailField;
        showNotification('Please enter a valid email address', 'error');
        isValid = false;
    }
    
    // Validate items
    const itemRows = itemsContainer.querySelectorAll('.item-row');
    let hasValidItems = false;
    
    itemRows.forEach(row => {
        const description = row.querySelector('input[name="itemDescription[]"]').value.trim();
        const rate = parseFloat(row.querySelector('input[name="rate[]"]').value) || 0;
        
        if (description && rate > 0) {
            hasValidItems = true;
        }
    });
    
    if (!hasValidItems) {
        showNotification('Please add at least one valid item with description and rate', 'error');
        isValid = false;
    }
    
    // Validate dates
    const documentDate = new Date(document.getElementById('documentDate').value);
    const dateField = currentDocumentType === 'quotation' ? 'validTillDate' : 'dueDate';
    const comparisonDate = new Date(document.getElementById(dateField).value);
    
    if (comparisonDate <= documentDate) {
        const field = document.getElementById(dateField);
        field.style.borderColor = '#dc2626';
        field.style.boxShadow = '0 0 0 3px rgba(220, 38, 38, 0.1)';
        if (!firstInvalidField) firstInvalidField = field;
        
        const fieldLabel = currentDocumentType === 'quotation' ? 'Valid till date' : 'Due date';
        showNotification(`${fieldLabel} must be after ${currentDocumentType} date`, 'error');
        isValid = false;
    }
    
    if (!isValid) {
        if (firstInvalidField) {
            firstInvalidField.focus();
            firstInvalidField.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        showNotification('Please fill in all required fields correctly', 'error');
    }
    
    return isValid;
}

/**
 * Get form data as object
 */
function getFormData() {
    const data = {
        id: Date.now().toString(),
        documentType: currentDocumentType,
        documentNo: document.getElementById('documentNo').value,
        documentDate: document.getElementById('documentDate').value,
        companyName: document.getElementById('companyName').value.trim(),
        companyAddress: document.getElementById('companyAddress').value.trim(),
        clientName: document.getElementById('clientName').value.trim(),
        clientEmail: document.getElementById('clientEmail').value.trim(),
        clientPhone: document.getElementById('clientPhone').value.trim(),
        clientLocation: document.getElementById('clientLocation').value.trim(),
        termsConditions: document.getElementById('termsConditions').value.trim(),
        items: [],
        timestamp: new Date().toISOString()
    };
    
    // Add document type specific fields
    if (currentDocumentType === 'quotation') {
        data.validTillDate = document.getElementById('validTillDate').value;
    } else {
        data.dueDate = document.getElementById('dueDate').value;
        data.paymentStatus = document.getElementById('paymentStatus')?.value || 'unpaid';
        data.paymentTerms = document.getElementById('paymentTerms')?.value || 'immediate';
        data.paymentMethod = document.getElementById('paymentMethod')?.value || 'bank';
        data.advanceAmount = parseFloat(document.getElementById('advanceAmount')?.value || 0);
    }
    
    // Get items data
    const itemRows = itemsContainer.querySelectorAll('.item-row');
    itemRows.forEach(row => {
        const description = row.querySelector('input[name="itemDescription[]"]').value.trim();
        const gstRate = parseFloat(row.querySelector('select[name="gstRate[]"]').value) || 0;
        const quantity = parseFloat(row.querySelector('input[name="quantity[]"]').value) || 0;
        const rate = parseFloat(row.querySelector('input[name="rate[]"]').value) || 0;
        const amount = parseFloat(row.querySelector('input[name="amount[]"]').value) || 0;
        const cgst = parseFloat(row.querySelector('input[name="cgst[]"]').value) || 0;
        const sgst = parseFloat(row.querySelector('input[name="sgst[]"]').value) || 0;
        const total = parseFloat(row.querySelector('input[name="total[]"]').value) || 0;
        
        if (description && rate > 0) {
            data.items.push({
                id: Date.now().toString() + Math.random(),
                description,
                gstRate,
                quantity,
                rate,
                amount,
                cgst,
                sgst,
                total
            });
        }
    });
    
    // Calculate totals
    data.subtotal = data.items.reduce((sum, item) => sum + item.amount, 0);
    data.totalCGST = data.items.reduce((sum, item) => sum + item.cgst, 0);
    data.totalSGST = data.items.reduce((sum, item) => sum + item.sgst, 0);
    data.grandTotal = data.subtotal + data.totalCGST + data.totalSGST;
    
    return data;
}

/**
 * Check for duplicates
 */
function checkForDuplicates(formData) {
    const currentData = currentDocumentType === 'quotation' ? allQuotations : allInvoices;
    return currentData.some(doc => 
        doc.clientEmail.toLowerCase() === formData.clientEmail.toLowerCase() &&
        doc.companyName.toLowerCase() === formData.companyName.toLowerCase()
    );
}

/**
 * Generate document preview
 */
function generateDocumentPreview(data) {
    const formatDate = (dateString) => {
        if (!dateString) return 'Not specified';
        const date = new Date(dateString);
        return date.toLocaleDateString('en-IN', {
            day: '2-digit',
            month: 'short',
            year: 'numeric'
        });
    };
    
    const docTypeName = (data.documentType || currentDocumentType).charAt(0).toUpperCase() + (data.documentType || currentDocumentType).slice(1);
    const isInvoice = (data.documentType || currentDocumentType) === 'invoice';
    
    let itemsHTML = '';
    data.items.forEach((item, index) => {
        itemsHTML += `
            <tr>
                <td>${index + 1}</td>
                <td>${item.description}</td>
                <td>${item.gstRate}%</td>
                <td>${item.quantity}</td>
                <td>₹${item.rate.toFixed(2)}</td>
                <td>₹${item.amount.toFixed(2)}</td>
                <td>₹${item.cgst.toFixed(2)}</td>
                <td>₹${item.sgst.toFixed(2)}</td>
                <td>₹${item.total.toFixed(2)}</td>
            </tr>
        `;
    });
    
    // Payment status badge for invoices
    const paymentStatusBadge = isInvoice ? `
        <div class="payment-status ${data.paymentStatus || 'unpaid'}">
            <i class="fas fa-circle"></i>
            ${(data.paymentStatus || 'unpaid').charAt(0).toUpperCase() + (data.paymentStatus || 'unpaid').slice(1)}
        </div>
    ` : '';
    
    // Calculate amount due for invoices
    const amountDue = isInvoice ? data.grandTotal - (data.advanceAmount || 0) : data.grandTotal;
    
    const html = `
        <div class="document-header">
            <div class="company-info">
                <div class="company-logo">
                    <img src="./logo.png" alt="Company Logo" onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';" style="width: 100%;">
                        <div style="display: none; width: 100%; height: 100%; align-items: center; justify-content: center; background: linear-gradient(135deg, #e5e9f0ff, #ebedf3ff); border-radius: 15px; color: white; font-size: 2rem;">
                        </div>
                </div>
                <div>
                    <h1>${data.companyName}</h1>
                    <p>${data.companyAddress.replace(/\n/g, '<br>')}</p>
                </div>
            </div>
            <div class="document-details">
                <h2>${docTypeName}</h2>
                ${paymentStatusBadge}
                <p><strong>${docTypeName} No:</strong> ${data.documentNo}</p>
                <p><strong>${docTypeName} Date:</strong> ${formatDate(data.documentDate)}</p>
                ${isInvoice ? 
                    `<p><strong>Due Date:</strong> ${formatDate(data.dueDate || data.validTillDate)}</p>` :
                    `<p><strong>Valid Till:</strong> ${formatDate(data.validTillDate)}</p>`
                }
            </div>
        </div>
        
        <div class="client-section">
            <div class="client-info">
                <h3>${docTypeName} From</h3>
                <p><strong>${data.companyName}</strong></p>
                <p>${data.companyAddress.replace(/\n/g, '<br>')}</p>
            </div>
            <div class="client-info">
                <h3>${docTypeName} To</h3>
                <p><strong>${data.clientName}</strong></p>
                <p><i class="fas fa-envelope"></i> ${data.clientEmail}</p>
                ${data.clientPhone ? `<p><i class="fas fa-phone"></i> ${data.clientPhone}</p>` : ''}
                ${data.clientLocation ? `<p><i class="fas fa-map-marker-alt"></i> ${data.clientLocation}</p>` : ''}
            </div>
        </div>
        
        <div class="items-table-wrapper">
            <table class="items-table">
                <thead>
                    <tr>
                        <th>S.No</th>
                        <th>Description</th>
                        <th>GST Rate</th>
                        <th>Quantity</th>
                        <th>Rate</th>
                        <th>Amount</th>
                        <th>CGST</th>
                        <th>SGST</th>
                        <th>Total</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsHTML}
                </tbody>
            </table>
        </div>
        
        <div class="totals-summary">
            <table class="totals-table">
                <tr>
                    <td>Subtotal</td>
                    <td>₹${data.subtotal.toFixed(2)}</td>
                </tr>
                <tr>
                    <td>CGST</td>
                    <td>₹${data.totalCGST.toFixed(2)}</td>
                </tr>
                <tr>
                    <td>SGST</td>
                    <td>₹${data.totalSGST.toFixed(2)}</td>
                </tr>
                ${isInvoice && data.advanceAmount > 0 ? `
                <tr>
                    <td>Advance Received</td>
                    <td>₹${data.advanceAmount.toFixed(2)}</td>
                </tr>` : ''}
                <tr class="grand-total">
                    <td><strong>${isInvoice ? 'Amount Due (INR)' : 'Grand Total (INR)'}</strong></td>
                    <td><strong>₹${amountDue.toFixed(2)}</strong></td>
                </tr>
            </table>
        </div>
        
        ${isInvoice && data.paymentTerms ? `
        <div style="margin-bottom: 30px; padding: 20px; background: #f8fafc; border-radius: 10px; border-left: 4px solid #3b82f6;">
            <h4 style="color: #3b82f6; margin-bottom: 10px;">Payment Information</h4>
            <p><strong>Payment Terms:</strong> ${data.paymentTerms.replace(/(\d+)days/, '$1 days').replace(/immediate/, 'Due Immediately')}</p>
            <p><strong>Payment Method:</strong> ${data.paymentMethod.charAt(0).toUpperCase() + data.paymentMethod.slice(1)}</p>
        </div>
        ` : ''}
        
        <div class="signature-section">
            <div class="signature-line"></div>
            <p><strong>Authorized Signatory</strong></p>
        </div>
        
        <div class="terms-section">
            <h3>Terms and Conditions</h3>
            <p>${data.termsConditions.replace(/\n/g, '<br>')}</p>
        </div>
    `;
    
    documentPreview.innerHTML = html;
}

/**
 * Show preview section
 */
function showPreview() {
    if (!currentDocumentData) {
        showNotification(`Please generate a ${currentDocumentType} first`, 'warning');
        return;
    }
    
    document.getElementById('formSection').style.display = 'none';
    previewSection.style.display = 'block';
    
    // Close mobile menu if open
    if (sidebar.classList.contains('active')) {
        sidebar.classList.remove('active');
    }
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/**
 * Hide preview section
 */
function hidePreview() {
    previewSection.style.display = 'none';
    document.getElementById('formSection').style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/**
 * Edit document
 */
function editDocument() {
    hidePreview();
}

/**
 * Print document
 */
function printDocument() {
    setTimeout(() => {
        window.print();
    }, 100);
}

/**
 * Clear form
 */
function clearForm(showConfirm = true) {
    if (showConfirm && !confirm('Are you sure you want to clear the form? All unsaved data will be lost.')) {
        return;
    }
    
    form.reset();
    
    // Reset items to single row
    itemsContainer.innerHTML = `
        <div class="item-row hover-lift" data-item="1">
            <div class="item-grid">
                <div class="form-group item-description">
                    <label>Item Description</label>
                    <input type="text" name="itemDescription[]" placeholder="Enter item description" required>
                </div>
                <div class="form-group">
                    <label>GST Rate (%)</label>
                    <select name="gstRate[]" class="gst-rate">
                        <option value="0">0%</option>
                        <option value="5">5%</option>
                        <option value="12">12%</option>
                        <option value="18" selected>18%</option>
                        <option value="28">28%</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Quantity</label>
                    <input type="number" name="quantity[]" min="1" value="1" class="quantity" required>
                </div>
                <div class="form-group">
                    <label>Rate (₹)</label>
                    <input type="number" name="rate[]" min="0" step="0.01" class="rate" placeholder="0.00" required>
                </div>
                <div class="form-group">
                    <label>Amount (₹)</label>
                    <input type="number" name="amount[]" class="amount" readonly>
                </div>
                <div class="form-group">
                    <label>CGST (₹)</label>
                    <input type="number" name="cgst[]" class="cgst" readonly>
                </div>
                <div class="form-group">
                    <label>SGST (₹)</label>
                    <input type="number" name="sgst[]" class="sgst" readonly>
                </div>
                <div class="form-group">
                    <label>Total (₹)</label>
                    <input type="number" name="total[]" class="item-total" readonly>
                </div>
                <div class="form-group item-actions">
                    <label>&nbsp;</label>
                    <button type="button" class="btn btn-danger btn-sm remove-item">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </div>
        </div>
    `;
    
    itemCounter = 1;
    currentDocumentData = null;
    generateDocumentNumber();
    setDefaultDates();
    calculateTotals();
    hidePreview();
    localStorage.removeItem(FORM_DRAFT_KEY);
    setStatus('ready', 'Ready');
    
    if (showConfirm) {
        showNotification('Form cleared successfully', 'info');
    }
    
    // Reset field styles
    const allFields = form.querySelectorAll('input, select, textarea');
    allFields.forEach(field => {
        field.style.borderColor = '#e5e7eb';
        field.style.boxShadow = 'none';
    });
}

/**
 * New document
 */
function newDocument() {
    clearForm();
}

/**
 * Clear all data
 */
function clearAllData() {
    const docTypeName = currentDocumentType.charAt(0).toUpperCase() + currentDocumentType.slice(1);
    
    if (confirm(`Are you sure you want to delete all ${currentDocumentType} data? This action cannot be undone.`)) {
        try {
            if (currentDocumentType === 'quotation') {
                allQuotations = [];
                localStorage.removeItem(QUOTATIONS_KEY);
            } else {
                allInvoices = [];
                localStorage.removeItem(INVOICES_KEY);
            }
            
            localStorage.removeItem(FORM_DRAFT_KEY);
            updateSidebarStats();
            updateRecentList();
            showNotification(`All ${currentDocumentType} data cleared successfully`, 'success');
        } catch (error) {
            console.error('Error clearing data:', error);
            showNotification('Error clearing data', 'error');
        }
    }
}

/**
 * Update sidebar statistics
 */
function updateSidebarStats() {
    const currentData = currentDocumentType === 'quotation' ? allQuotations : allInvoices;
    totalDocuments.textContent = currentData.length;
    const total = currentData.reduce((sum, doc) => sum + (doc.grandTotal || 0), 0);
    totalAmount.textContent = `₹${total.toFixed(2)}`;
}

/**
 * Update recent documents list
 */
function updateRecentList() {
    const currentData = currentDocumentType === 'quotation' ? allQuotations : allInvoices;
    
    if (currentData.length === 0) {
        recentList.innerHTML = '<div class="no-data">No documents yet</div>';
        return;
    }
    
    const recent = currentData.slice(-5).reverse();
    recentList.innerHTML = recent.map(doc => `
        <div class="recent-item hover-lift" onclick="viewDocument('${doc.id}')">
            <div class="recent-info">
                <div class="recent-title">${doc.documentNo}</div>
                <div class="recent-client">${doc.clientName}</div>
                <div class="recent-amount">₹${doc.grandTotal.toFixed(2)}</div>
            </div>
        </div>
    `).join('');
}

/**
 * Handle search
 */
function handleSearch() {
    const searchTerm = searchInput.value.toLowerCase().trim();
    if (!searchTerm) {
        searchResults.innerHTML = '';
        return;
    }
    
    const currentData = currentDocumentType === 'quotation' ? allQuotations : allInvoices;
    const filtered = currentData.filter(doc => 
        doc.documentNo.toLowerCase().includes(searchTerm) ||
        doc.clientName.toLowerCase().includes(searchTerm) ||
        doc.clientEmail.toLowerCase().includes(searchTerm) ||
        doc.companyName.toLowerCase().includes(searchTerm)
    );
    
    if (filtered.length === 0) {
        searchResults.innerHTML = '<div class="no-data">No results found</div>';
        return;
    }
    
    searchResults.innerHTML = filtered.map(doc => `
        <div class="search-item hover-lift" onclick="viewDocument('${doc.id}')">
            <div class="search-info">
                <div class="search-title">${doc.documentNo}</div>
                <div class="search-client">${doc.clientName}</div>
                <div class="search-amount">₹${doc.grandTotal.toFixed(2)}</div>
            </div>
        </div>
    `).join('');
}

/**
 * View specific document
 */
window.viewDocument = function(id) {
    const allData = [...allQuotations, ...allInvoices];
    const document = allData.find(doc => doc.id === id);
    
    if (document) {
        // Switch to correct document type if necessary
        if (document.documentType !== currentDocumentType) {
            currentDocumentType = document.documentType;
            tabBtns.forEach(btn => btn.classList.remove('active'));
            document.querySelector(`[data-type="${currentDocumentType}"]`).classList.add('active');
            updateUIForDocumentType();
            updateSidebarStats();
            updateRecentList();
        }
        
        currentDocumentData = document;
        generateDocumentPreview(document);
        showPreview();
        
        if (sidebar.classList.contains('active')) {
            sidebar.classList.remove('active');
        }
    }
};

/**
 * Save form draft to localStorage
 */
function saveDraftToLocal() {
    try {
        if (!validateFormBasic()) return;
        
        const formData = getFormData();
        localStorage.setItem(FORM_DRAFT_KEY, JSON.stringify(formData));
        setStatus('draft', 'Draft saved');
        
        setTimeout(() => {
            if (statusIndicator.classList.contains('draft')) {
                setStatus('ready', 'Ready');
            }
        }, 2000);
    } catch (error) {
        console.error('Error saving draft:', error);
    }
}

/**
 * Load form draft
 */
function loadFormDraft(draftData) {
    try {
        // Switch to correct document type
        if (draftData.documentType && draftData.documentType !== currentDocumentType) {
            currentDocumentType = draftData.documentType;
            tabBtns.forEach(btn => btn.classList.remove('active'));
            document.querySelector(`[data-type="${currentDocumentType}"]`).classList.add('active');
            updateUIForDocumentType();
        }
        
        // Restore basic form fields
        const fieldMappings = {
            'documentDate': 'documentDate',
            'validTillDate': 'validTillDate',
            'dueDate': 'dueDate',
            'companyName': 'companyName',
            'companyAddress': 'companyAddress',
            'clientName': 'clientName',
            'clientEmail': 'clientEmail',
            'clientPhone': 'clientPhone',
            'clientLocation': 'clientLocation',
            'termsConditions': 'termsConditions',
            'paymentStatus': 'paymentStatus',
            'paymentTerms': 'paymentTerms',
            'paymentMethod': 'paymentMethod',
            'advanceAmount': 'advanceAmount'
        };
        
        Object.entries(fieldMappings).forEach(([formField, dataField]) => {
            const element = document.getElementById(formField);
            if (element && draftData[dataField]) {
                element.value = draftData[dataField];
            }
        });
        
        // Restore items if available
        if (draftData.items && draftData.items.length > 0) {
            restoreFormItems(draftData.items);
        }
        
        calculateTotals();
        showNotification('Draft data restored', 'info');
    } catch (error) {
        console.error('Error loading form draft:', error);
    }
}

/**
 * Restore form items from draft data
 */
function restoreFormItems(items) {
    // Clear existing items
    itemsContainer.innerHTML = '';
    itemCounter = 0;
    
    // Add items from draft
    items.forEach((item, index) => {
        addNewItem();
        const itemRow = itemsContainer.lastElementChild;
        
        itemRow.querySelector('input[name="itemDescription[]"]').value = item.description || '';
        itemRow.querySelector('select[name="gstRate[]"]').value = item.gstRate || 18;
        itemRow.querySelector('input[name="quantity[]"]').value = item.quantity || 1;
        itemRow.querySelector('input[name="rate[]"]').value = item.rate || 0;
        
        // Trigger calculation
        itemRow.querySelector('input[name="rate[]"]').dispatchEvent(new Event('input', { bubbles: true }));
    });
}

/**
 * Basic form validation for draft saving
 */
function validateFormBasic() {
    const companyName = document.getElementById('companyName').value.trim();
    const clientName = document.getElementById('clientName').value.trim();
    return companyName || clientName;
}

/**
 * Set status indicator
 */
function setStatus(status, message) {
    statusIndicator.className = `status-indicator ${status}`;
    statusIndicator.querySelector('span').textContent = message;
    
    const icon = statusIndicator.querySelector('i');
    switch (status) {
        case 'ready':
            icon.className = 'fas fa-circle';
            break;
        case 'draft':
            icon.className = 'fas fa-edit';
            break;
        case 'saving':
            icon.className = 'fas fa-spinner fa-spin';
            break;
        case 'saved':
            icon.className = 'fas fa-check-circle';
            break;
        case 'error':
            icon.className = 'fas fa-exclamation-circle';
            break;
        case 'duplicate':
            icon.className = 'fas fa-exclamation-triangle';
            break;
    }
    
    // Reset status after delay
    if (status !== 'ready' && status !== 'error') {
        setTimeout(() => {
            if (statusIndicator.classList.contains(status)) {
                setStatus('ready', 'Ready');
            }
        }, 3000);
    }
}

/**
 * Show notification
 */
function showNotification(message, type = 'info') {
    const notificationIcon = notification.querySelector('.notification-icon');
    const notificationMessage = notification.querySelector('.notification-message');
    
    notification.className = `notification ${type}`;
    notificationMessage.textContent = message;
    
    switch (type) {
        case 'success':
            notificationIcon.className = 'notification-icon fas fa-check-circle';
            break;
        case 'error':
            notificationIcon.className = 'notification-icon fas fa-exclamation-circle';
            break;
        case 'warning':
            notificationIcon.className = 'notification-icon fas fa-exclamation-triangle';
            break;
        case 'info':
            notificationIcon.className = 'notification-icon fas fa-info-circle';
            break;
    }
    
    notification.classList.add('show');
    
    setTimeout(() => {
        notification.classList.remove('show');
    }, 5000);
}

/**
 * Debounce function
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}