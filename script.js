let jsonData = [];

document.getElementById("excelFile").addEventListener("change", handleFileUpload);
document.getElementById("submitButton").addEventListener("click", () => {
    displayData(jsonData);
});
document.getElementById("filterCategory").addEventListener("change", filterData);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: "array"
            });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            jsonData = XLSX.utils.sheet_to_json(sheet, {
                header: 1
            });
        };
        reader.readAsArrayBuffer(file);
    }
}

function displayData(data) {
    const tableBody = document.getElementById("caseTable");
    tableBody.innerHTML = "";

    data.forEach((row, index) => {
        if (index === 0 || row.length < 10) return;

        const [no, tanggalRaw, sidangKeliling, ruangan, nomorPerkara, agenda, penggugat, tergugat,
            majelisHakim, paniteraPengganti
        ] = row;

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

function addRowToTable(caseData) {
    const tableBody = document.getElementById("caseTable");
    const row = document.createElement("tr");
    row.innerHTML = `
        <td>${caseData.no}</td>
        <td>${caseData.tanggal}</td>
        <td>${caseData.ruangan}</td>
        <td style="font-weight: bold; background-color: #FDFEEF"">${caseData.nomorPerkara}</td>
        <td>${caseData.agenda}</td>
        <td style="font-weight: bold; background-color: #FDFEEF">${caseData.majelisHakim}</td>
        <td>${caseData.paniteraPengganti}</td>
        <td class="tergugat-cell">${caseData.tergugat}</td>
        <td style="font-weight: bold; background-color: #F0F2FF ">${caseData.penggugat}</td>
        <td>
            <select class="form-control mb-2" onchange="prepareWhatsAppNotification(this, '${caseData.nomorPerkara}', '${caseData.tanggal}', '${caseData.ruangan}', '${caseData.agenda}', '${caseData.penggugat}', '${caseData.tergugat}', '${caseData.majelisHakim}', '${caseData.paniteraPengganti}')">
                <option value="">Pilih Nomor</option>
                <option value="+6282181361433">+6282181361433 (Apri)</option>
                <option value="+6289876543210">+6289876543210</option>
                <option value="+6281122334455">+6281122334455</option>
            </select>
            <button class="btn btn-notification mt-2" onclick="sendWhatsAppNotification(this)">Kirim Notifikasi</button>
        </td>
    `;
    tableBody.appendChild(row);
}

function filterData() {
    const category = document.getElementById("filterCategory").value.toLowerCase();
    const filteredData = category ?
        jsonData.filter(row => row[4] && row[4].toLowerCase().includes(category)) :
        jsonData;

    displayData(filteredData);
}

function prepareWhatsAppNotification(selectElement, nomorPerkara, tanggal, ruangan, agenda, penggugat, tergugat,
    majelisHakim, paniteraPengganti) {
    const selectedPhone = selectElement.value;
    if (!selectedPhone) return;

    const messageTemplate =
        `PEMBERITAHUAN PERSIDANGAN AKAN DI MULAI!!\nDi Pengailan Negeri Gunung Sugih Kelas IB \nTanggal: ${tanggal}\n\nDetail Persidangan :\nNomor Perkara: ${nomorPerkara}\nRuangan: ${ruangan}\nAgenda: ${agenda}\nPenggugat: ${penggugat}\nTergugat: ${tergugat}\nMajelis Hakim: ${majelisHakim}\nPanitera Pengganti: ${paniteraPengganti}\n\nMohon untuk hadir tepat waktu. Terima kasih.`;
    selectElement.closest("td").querySelector('.btn-notification').dataset.message =
        messageTemplate;
}

function sendWhatsAppNotification(buttonElement) {
    const selectedPhone = buttonElement.closest("td").querySelector("select").value;
    const message = buttonElement.dataset.message;

    if (!selectedPhone || !message) {
        alert("Pilih nomor telepon terlebih dahulu.");
        return;
    }

    const whatsappLink =
        `https://api.whatsapp.com/send?phone=${selectedPhone}&text=${encodeURIComponent(message)}`;
    window.open(whatsappLink, '_blank');
}