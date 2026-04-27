// Staff Retirement Management System Logic
// Using global XLSX from script tag in index.html

document.addEventListener('DOMContentLoaded', () => {
    console.log('Retirement System Initialized');

    // State Management
    let staffData = [];
    let retirees = [];
    const currentYear = new Date().getFullYear();

    // DOM Elements
    const sections = {
        upload: document.getElementById('section-upload'),
        retirees: document.getElementById('section-retirees'),
        reports: document.getElementById('section-reports')
    };

    const navBtns = {
        upload: document.getElementById('btn-upload'),
        retirees: document.getElementById('btn-retirees'),
        reports: document.getElementById('btn-reports')
    };

    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');
    const dropZone = document.getElementById('drop-zone');
    const retireesBody = document.getElementById('retirees-body');
    const toast = document.getElementById('toast');

    // Initialize UI
    const yearEl = document.getElementById('current-year');
    if (yearEl) yearEl.textContent = currentYear;

    // Navigation Logic
    Object.keys(navBtns).forEach(key => {
        const btn = navBtns[key];
        if (btn) {
            btn.addEventListener('click', () => {
                console.log(`Switching to section: ${key}`);
                // Update active nav button
                Object.values(navBtns).forEach(b => b && b.classList.remove('active'));
                btn.classList.add('active');

                // Show active section
                Object.values(sections).forEach(s => s && s.classList.remove('active'));
                if (sections[key]) sections[key].classList.add('active');
            });
        }
    });

    // File Upload Logic
    if (browseBtn && fileInput) {
        browseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            console.log('Browse button clicked');
            fileInput.click();
        });
        
        fileInput.addEventListener('change', handleFileSelect);
    }

    if (dropZone) {
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-over');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('drag-over');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                processFile(files[0]);
            }
        });
    }

    function handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            processFile(file);
        }
    }

    function processFile(file) {
        console.log('Processing file:', file.name);
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                processStaffData(jsonData);
                showToast('Records uploaded and processed successfully!');
                
                // Switch to retirees view
                if (navBtns.retirees) navBtns.retirees.click();
            } catch (err) {
                console.error('Error processing file:', err);
                showToast('Error: Could not process file. Check format.');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function processStaffData(data) {
        staffData = data;
        retirees = [];

        data.forEach(staff => {
            const dob = parseDate(staff['Date of Birth'] || staff['DOB'] || staff['dob']);
            const dofa = parseDate(staff['Date of First Appointment'] || staff['DOFA'] || staff['dofa']);
            const name = staff['Staff Name'] || staff['Name'] || staff['name'];
            const id = staff['Staff ID'] || staff['Unique ID'] || staff['ID'] || staff['id'];

            if (!dob || !dofa) return;

            // Retirement logic
            const age60Date = new Date(dob);
            age60Date.setFullYear(age60Date.getFullYear() + 60);

            const service35Date = new Date(dofa);
            service35Date.setFullYear(service35Date.getFullYear() + 35);

            // Retirement date is the EARLIER of the two
            let retirementDate, reason;
            if (age60Date < service35Date) {
                retirementDate = age60Date;
                reason = 'Reached age 60';
            } else {
                retirementDate = service35Date;
                reason = '35 years of service';
            }

            // Check if retiring in CURRENT year or earlier
            if (retirementDate.getFullYear() <= currentYear) {
                retirees.push({
                    id,
                    name,
                    dob: dob.toLocaleDateString(),
                    dofa: dofa.toLocaleDateString(),
                    retirementDate: retirementDate.toLocaleDateString(),
                    rawRetirementDate: retirementDate,
                    reason
                });
            }
        });

        updateUI();
    }

    function parseDate(dateInput) {
        if (!dateInput) return null;
        
        // Handle Excel serial dates
        if (typeof dateInput === 'number') {
            return new Date((dateInput - 25569) * 86400 * 1000);
        }
        
        const d = new Date(dateInput);
        return isNaN(d.getTime()) ? null : d;
    }

    function updateUI() {
        // Update Stats
        const elTotal = document.getElementById('total-processed');
        const elDue = document.getElementById('due-retirement');
        const elAge = document.getElementById('retire-age');
        const elService = document.getElementById('retire-service');

        if (elTotal) elTotal.textContent = staffData.length;
        if (elDue) elDue.textContent = retirees.length;
        if (elAge) elAge.textContent = retirees.filter(r => r.reason.includes('age')).length;
        if (elService) elService.textContent = retirees.filter(r => r.reason.includes('service')).length;

        // Update Table
        if (retireesBody) {
            if (retirees.length === 0) {
                retireesBody.innerHTML = `
                    <tr class="empty-state">
                        <td colspan="7">No retirees found for the current criteria.</td>
                    </tr>
                `;
            } else {
                // Sort by retirement date
                retirees.sort((a, b) => a.rawRetirementDate - b.rawRetirementDate);
                
                retireesBody.innerHTML = retirees.map(r => `
                    <tr>
                        <td>${r.id || 'N/A'}</td>
                        <td><strong>${r.name}</strong></td>
                        <td>${r.dob}</td>
                        <td>${r.dofa}</td>
                        <td class="highlight">${r.retirementDate}</td>
                        <td><span class="badge ${r.reason.includes('age') ? 'age' : 'service'}">${r.reason}</span></td>
                        <td>
                            <button class="action-icon" title="Send Reminder">
                                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 17a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V9.5C2 7 4 5 6.5 5H18c2.5 0 4 2 4 4.5V17z"></path><polyline points="15,7 12,10 9,7"></polyline></svg>
                            </button>
                        </td>
                    </tr>
                `).join('');
            }
        }
    }

    function showToast(message) {
        if (!toast) return;
        toast.textContent = message;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    // Print logic
    const printBtn = document.getElementById('print-btn');
    if (printBtn) {
        printBtn.addEventListener('click', () => {
            window.print();
        });
    }

    // Export logic (simple CSV)
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', () => {
            if (retirees.length === 0) return;
            
            const headers = ['Staff ID', 'Name', 'DOB', 'DOFA', 'Retirement Date', 'Reason'];
            const rows = retirees.map(r => [r.id, r.name, r.dob, r.dofa, r.retirementDate, r.reason]);
            
            let csvContent = "data:text/csv;charset=utf-8," 
                + headers.join(",") + "\n"
                + rows.map(e => e.join(",")).join("\n");
                
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", `retirees_${currentYear}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    }
});
