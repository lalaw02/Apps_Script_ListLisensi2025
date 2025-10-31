function kirimTagihanReminder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const adminEmail = 
  "admin@gmail.com"; 

  let daftarTagihan = [];

  for (let i = 1; i < data.length; i++) {
    const productName = data[i][3];   // kolom D: Product Name
    const endDate = data[i][9];       // kolom J: End Date
    const minusDays = data[i][11];    // kolom L: Minus Days
    const pic = data[i][13];          // kolom N: PIC Unit Kerja

    if (typeof minusDays === "number" && [30, 7, 1, 0].includes(minusDays)) {
      daftarTagihan.push({
        product: productName,
        endDate: new Date(endDate),
        minus: minusDays,
        pic: pic
      });
    }
  }

  daftarTagihan.sort((a, b) => a.minus - b.minus);

  if (daftarTagihan.length === 0) {
    Logger.log("Tidak ada tagihan H-30, H-7, H-1, atau H-0 hari ini.");
    return;
  }

  // Buat isi email HTML
  let htmlBody = `
  <p>Yth. Bapak/Ibu <b>Deputi Bidang Infrastruktur dan Operasional TI</b>,</p>
  <p><b>Kindly Reminder,</b></p>
  <p>Berikut daftar lisensi aplikasi yang mendekati end date:</p>
  <ol style="margin:0; padding-left:18px;">
  `;

  daftarTagihan.forEach((item) => {
    const tglFormat = Utilities.formatDate(item.endDate, "Asia/Jakarta", "dd MMMM yyyy");
    let status = "";

    if (item.minus === 0) {
      status = "<b>Hari ini adalah end date</b>";
    } else if (item.minus === 1) {
      status = "<b>Besok adalah end date</b>";
    } else {
      status = <b>${item.minus} hari lagi menuju end date</b>;
    }

    htmlBody += `
      <li style="margin-bottom:6px;">
        ${status} : <b>${item.product}</b> (${item.pic}) – ${tglFormat}
      </li>
    `;
  });

  htmlBody += `
  </ol>
  <p>Terima kasih.</p>
  `;

  MailApp.sendEmail({
    to: adminEmail,
    subject: "Reminder Lisensi Aplikasi Jatuh Tempo",
    htmlBody: htmlBody
  });
}
