let jsonData = [];

document.getElementById("excelFile").addEventListener("change", handleFileUpload);
document.getElementById("submitButton").addEventListener("click", () => {
    const filteredData = getFilteredData();
    renderTable(filteredData);
});
document.getElementById("filterCategory").addEventListener("input", () => {
    const filteredData = getFilteredData();
    renderTable(filteredData);
});
document.getElementById("filterAgenda").addEventListener("input", () => {
    const filteredData = getFilteredData();
    renderTable(filteredData);
});
document.getElementById("filterMajelis").addEventListener("input", () => {
    const filteredData = getFilteredData();
    renderTable(filteredData);
});
document.getElementById("filterPanitera").addEventListener("input", () => {
    const filteredData = getFilteredData();
    renderTable(filteredData);
});

// Handle file upload and convert Excel to JSON
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            renderTable(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }
}

// Render table with filtered data
function renderTable(data) {
    const tableBody = document.getElementById("caseTable");
    tableBody.innerHTML = "";

    if (data.length === 0 || (data.length === 1 && data[0].length < 10)) {
        tableBody.innerHTML = '<tr id="noDataMessage"><td colspan="10" class="text-center text-muted">Tidak ada data</td></tr>';
        return;
    }

    data.forEach((row, index) => {
        if (index === 0 || row.length < 10) return;  // Skip header row or incomplete rows

        const [no, tanggalRaw, sidangKeliling, ruangan, nomorPerkara, agenda, penggugat, tergugat,
            majelisHakim, paniteraPengganti] = row;

        const tanggal = XLSX.SSF.format("dd-mmm-yy", tanggalRaw);

        const caseData = {
            no: index,
            tanggal: tanggal || "",
            ruangan: ruangan || "",
            nomorPerkara: nomorPerkara || "",
            agenda: agenda || "",
            penggugat: penggugat || "",
            tergugat: tergugat || "",
            majelisHakim: majelisHakim || "",
            paniteraPengganti: paniteraPengganti || ""
        };

        addRowToTable(caseData);
    });
}

// Add a row to the table
function addRowToTable(caseData) {
    const tableBody = document.getElementById("caseTable");
    const row = document.createElement("tr");
    row.innerHTML = `
        <td>${caseData.no}</td>
        <td>${caseData.tanggal}</td>
        <td>${caseData.ruangan}</td>
        <td style="font-weight: bold; background-color: #FDFEEF">${caseData.nomorPerkara}</td>
        <td>${caseData.agenda}</td>
        <td style="font-weight: bold; background-color: #FDFEEF">${caseData.majelisHakim}</td>
        <td style="font-weight: bold; background-color: #ECFFE3">${caseData.paniteraPengganti}</td>
        <td class="tergugat-cell" style="font-weight: bold; background-color: #F0F2FF">${caseData.tergugat}</td>
        <td>${caseData.penggugat}</td>
        <td>
    <select class="form-control mb-2" onchange="prepareWhatsAppNotification(this, '${caseData.nomorPerkara}', '${caseData.tanggal}', '${caseData.ruangan}', '${caseData.agenda}', '${caseData.penggugat}', '${caseData.tergugat}', '${caseData.majelisHakim}', '${caseData.paniteraPengganti}')">
        <option value="">Pilih</option>
        <option value="+6289876543210">Apri</option>
        <option value="+6282181361433">Gatot Kaca</option>
        <!-- Add more options here -->
    </select>
    <button class="btn btn-notification mt-2" onclick="sendWhatsAppNotification(this)">
        <i class="fas fa-paper-plane"></i> Kirim
    </button>
</td>
<td>
    <select class="form-control mb-2" onchange="prepareEmailNotification(this, '${caseData.nomorPerkara}', '${caseData.tanggal}', '${caseData.ruangan}', '${caseData.agenda}', '${caseData.penggugat}', '${caseData.tergugat}', '${caseData.majelisHakim}', '${caseData.paniteraPengganti}')">
        <option value="">Pilih</option>
        <option value="email1@example.com">Apri</option>
        <option value="email2@example.com">Gatot Kaca</option>
        <!-- Tambahkan lebih banyak alamat email di sini -->
    </select>
    <button class="btn btn-notification mt-2" onclick="sendEmailNotification(this)">
        <i class="fas fa-paper-plane"></i> Kirim
    </button>
</td>

    `;
    tableBody.appendChild(row);
}

// Get filtered data based on inputs
function getFilteredData() {
    const category = document.getElementById("filterCategory").value.toLowerCase();
    const agenda = document.getElementById("filterAgenda").value.toLowerCase();
    const majelis = document.getElementById("filterMajelis").value.toLowerCase();
    const panitera = document.getElementById("filterPanitera").value.toLowerCase();

    return jsonData.filter((row, index) => {
        if (index === 0 || row.length < 10) return false;  // Skip header row or incomplete rows

        const matchesCategory = category ? row[4]?.toLowerCase().includes(category) : true;
        const matchesAgenda = agenda ? row[5]?.toLowerCase().includes(agenda) : true;
        const matchesMajelis = majelis ? row[8]?.toLowerCase().includes(majelis) : true;
        const matchesPanitera = panitera ? row[9]?.toLowerCase().includes(panitera) : true;

        return matchesCategory && matchesAgenda && matchesMajelis && matchesPanitera;
    });
}

// Prepare WhatsApp notification message template
function prepareWhatsAppNotification(selectElement, nomorPerkara, tanggal, ruangan, agenda, penggugat, tergugat, majelisHakim, paniteraPengganti) {
    const selectedPhone = selectElement.value;
    if (!selectedPhone) return;

    const messageTemplate =
        `PEMBERITAHUAN PERSIDANGAN AKAN DI MULAI!!\nDi Pengadilan Negeri Gunung Sugih Kelas IB\nTanggal: ${tanggal}\n\nDetail Persidangan:\nNomor Perkara: ${nomorPerkara}\nRuangan: ${ruangan}\nAgenda: ${agenda}\nPenggugat: ${penggugat}\nTergugat: ${tergugat}\nMajelis Hakim: ${majelisHakim}\nPanitera Pengganti: ${paniteraPengganti}\n\nMohon untuk hadir tepat waktu. Terima kasih.`;
    selectElement.closest("td").querySelector('.btn-notification').dataset.message = messageTemplate;
}

// Send WhatsApp notification
function sendWhatsAppNotification(buttonElement) {
    const selectedPhone = buttonElement.closest("td").querySelector("select").value;
    const message = buttonElement.dataset.message;

    if (!selectedPhone || !message) {
        alert("Pilih nomor telepon terlebih dahulu.");
        return;
    }

    const whatsappLink = `https://api.whatsapp.com/send?phone=${selectedPhone}&text=${encodeURIComponent(message)}`;
    window.open(whatsappLink, '_blank');
}

// Search functionality
document.getElementById('searchBox').addEventListener('input', function () {
    const searchTerm = this.value.toLowerCase();
    const rows = document.querySelectorAll('#caseTable tr');

    rows.forEach(row => {
        const rowText = row.textContent.toLowerCase();
        row.style.display = rowText.includes(searchTerm) ? '' : 'none';
    });
});

// Download Word
function downloadWord() {
    const filteredData = getFilteredData();
    const tableRows = filteredData.map((row, index) => `
        <tr>
            <td>${index + 1}</td>
            <td>${XLSX.SSF.format("dd-mmm-yy", row[1]) || ""}</td>
            <td>${row[3] || ""}</td>
            <td>${row[4] || ""}</td>
            <td>${row[5] || ""}</td>
            <td>${row[8] || ""}</td>
            <td>${row[9] || ""}</td>
            <td>${row[7] || ""}</td>
            <td>${row[6] || ""}</td>
        </tr>
    `).join("");

    const category = document.getElementById("filterCategory").value || "Semua";
    const titleText = `Jadwal Sidang ${category.charAt(0).toUpperCase() + category.slice(1)}\nPengadilan Negeri Gunung Sugih - ${new Date().toLocaleDateString()}`;

    const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'>
        <head>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; }
                h2 { text-align: center; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { border: 1px solid #000; padding: 8px; text-align: center; }
                th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>
            <h2>${titleText}</h2>
            <table>
                <tr>
                    <th>No</th>
                    <th>Tanggal</th>
                    <th>Ruangan</th>
                    <th>Nomor Perkara</th>
                    <th>Agenda</th>
                    <th>Majelis Hakim</th>
                    <th>Panitera Pengganti</th>
                    <th>Tergugat</th>
                    <th>Penggugat</th>
                </tr>
                ${tableRows}
            </table>
        </body>
        </html>`;

    const blob = new Blob(['\ufeff', htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Jadwal_Sidang_${category}_${new Date().toLocaleDateString()}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// Download PDF
async function downloadPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    const filteredData = getFilteredData();
    const category = document.getElementById("filterCategory").value || "Semua";
    const titleText = `Jadwal Sidang ${category.charAt(0).toUpperCase() + category.slice(1)} - Pengadilan Negeri Gunung Sugih\nTanggal: ${new Date().toLocaleDateString()}`;
    
    doc.setFontSize(14);
    doc.text(titleText, 14, 15);

    const rows = filteredData.map((row, index) => [
        index + 1,
        XLSX.SSF.format("dd-mmm-yy", row[1]) || "",
        row[3] || "",
        row[4] || "",
        row[5] || "",
        row[8] || "",
        row[9] || "",
        row[7] || "",
        row[6] || ""
    ]);

    doc.autoTable({
        head: [['No', 'Tanggal', 'Ruangan', 'Nomor Perkara', 'Agenda', 'Majelis Hakim', 'Panitera Pengganti', 'Tergugat', 'Penggugat']],
        body: rows,
        startY: 25,
        theme: 'grid',
        headStyles: { fillColor: [230, 230, 250], textColor: 0 },
        styles: { fontSize: 10, cellPadding: 3 },
        columnStyles: { 0: { cellWidth: 10 } }
    });

    doc.save(`Jadwal_Sidang_${category}_${new Date().toLocaleDateString()}.pdf`);
}

// Add event listeners to download buttons
document.getElementById("downloadWord").addEventListener("click", downloadWord);
document.getElementById("downloadPDF").addEventListener("click", downloadPDF);


// Fungsi untuk mempersiapkan dan mengirim email
emailjs.init("YOUR_USER_ID");

function prepareEmailNotification(selectElement, nomorPerkara, tanggal, ruangan, agenda, penggugat, tergugat, majelisHakim, paniteraPengganti) {
    const selectedEmail = selectElement.value;
    if (!selectedEmail) {
        alert("Pilih alamat email terlebih dahulu.");
        return;
    }

    const subject = `Pemberitahuan Persidangan: ${nomorPerkara}`;
    const body = `
        <h3>PEMBERITAHUAN PERSIDANGAN AKAN DIMULAI!</h3>
        <p><strong>Pengadilan Negeri Gunung Sugih Kelas IB</strong></p>
        <p><strong>Tanggal Sidang:</strong> ${tanggal}</p>
        <p><strong>Ruangan:</strong> ${ruangan}</p>
        <p><strong>Agenda:</strong> ${agenda}</p>
        <p><strong>Penggugat:</strong> ${penggugat}</p>
        <p><strong>Tergugat:</strong> ${tergugat}</p>
        <p><strong>Majelis Hakim:</strong> ${majelisHakim}</p>
        <p><strong>Panitera Pengganti:</strong> ${paniteraPengganti}</p>
        <p>Mohon untuk hadir tepat waktu. Terima kasih.</p>
    `;

    // Mengirim email menggunakan EmailJS
    emailjs.send("YOUR_SERVICE_ID", "YOUR_TEMPLATE_ID", {
        to_email: selectedEmail,
        subject: subject,
        message: body,
    }).then(
        (response) => {
            alert("Notifikasi email berhasil dikirim!");
        },
        (error) => {
            alert("Gagal mengirim notifikasi email: " + error.text);
        }
    );
}
