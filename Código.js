function sendWeeklyEmailsToPerformersGrouped() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Summary");
  const resilienceSheet = ss.getSheetByName("Resilience");

  // Validación básica
  if (!sheet || !resilienceSheet) {
    throw new Error("No se pudo encontrar alguna de las hojas necesarias (Summary o Resilience).");
  }

  const data = sheet.getDataRange().getValues();
  const resilienceData = resilienceSheet.getDataRange().getValues();

  const userMessages = {};

  // Construcción del mapa: user -> TL mail
  const teamLeadMap = resilienceData.slice(1).reduce((map, row) => {
    const user = row[14]; // Col O
    const lead = row[16]; // Col Q
    if (user && lead) {
      const username = user.toString().trim().toLowerCase();
      const leadEmail = `${lead.toString().trim().toLowerCase()}@google.com`;
      map[username] = leadEmail;
    }
    return map;
  }, {});

  // Procesamiento de datos de Summary
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const usernameCell = row[10]; // Columna K

    if (
      usernameCell &&
      row[0] &&
      row[1] &&
      isFinite(row[2]) &&
      isFinite(row[3]) &&
      row[4] !== undefined
    ) {
      const difference = Number(row[2]) - Number(row[3]);
      const formattedDifference = difference.toFixed(2);
      const varianceRaw = Number(row[4]);
      const line = [row[0], row[1], formattedDifference, varianceRaw];

      const usernames = usernameCell.split('/').map(name => name.trim().toLowerCase());

      usernames.forEach(username => {
        if (!userMessages[username]) {
          userMessages[username] = [];
        }
        userMessages[username].push(line);
      });
    }
  }

  // Mails
  for (const username in userMessages) {
    const recipientEmail = `${username}@google.com`;
    const ccEmail = teamLeadMap[username] || null;

    const htmlMessage = buildEmailHtml(userMessages[username]);

    const emailOptions = {
      to: recipientEmail,
      subject: `Weekly GPH Data Monitoring Summary - ${username}`,
      htmlBody: htmlMessage
    };

    if (ccEmail) {
      emailOptions.cc = ccEmail;
    }

    MailApp.sendEmail(emailOptions);
    Logger.log(`Email enviado a: ${recipientEmail} ${ccEmail ? '(CC: ' + ccEmail + ')' : ''}`);
  }
}

// Formato mail HTML
function buildEmailHtml(rows) {
  const tableRows = rows.map(row => {
    const variancePercentage = (row[3] * 100).toFixed(2) + '%';
    const varianceValue = row[3] * 100;
    const bgColor = (varianceValue >= -10 && varianceValue <= 10) ? '#d4edda' : '#f8d7da';

    return `<tr>
      <td style="padding: 8px; border: 1px solid #ccc;">${row[0]}</td>
      <td style="padding: 8px; border: 1px solid #ccc;">${row[1]}</td>
      <td style="padding: 8px; border: 1px solid #ccc; text-align:right;">${row[2]}</td>
      <td style="padding: 8px; border: 1px solid #ccc; text-align:right; background-color:${bgColor};">${variancePercentage}</td>
    </tr>`;
  }).join("");

  return `
    <p style="font-family: Arial, sans-serif; font-size: 14px;">
      This is your weekly summary of your GPH Data Monitoring:
    </p>
    <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
      <thead>
        <tr style="background-color: #343a40; color: white;">
          <th style="padding: 8px; border: 1px solid #ccc;">DTP Name</th>
          <th style="padding: 8px; border: 1px solid #ccc;">Frequency</th>
          <th style="padding: 8px; border: 1px solid #ccc; text-align:right;">Remaining Use</th>
          <th style="padding: 8px; border: 1px solid #ccc; text-align:right;">Variance</th>
        </tr>
      </thead>
      <tbody>
        ${tableRows}
      </tbody>
    </table>
  `;
}