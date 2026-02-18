// 1. Create the Custom Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Email Tool')
    .addItem('Open Email Sender', 'showSidebar')
    .addToUi();
}

// 2. Display the Sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Email Sender')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 3. Helper: Define your available HTML Template files here
function getTemplateList() {
  // IMPORTANT: These names must match your HTML filenames exactly!
  return [
    "templatesatu", 
    "templatedua",
  ];
}

// 4. Helper: Get valid Gmail aliases
function getUserAliases() {
  const aliases = GmailApp.getAliases();
  const primary = Session.getActiveUser().getEmail();
  let allEmails = [primary];
  if (aliases && aliases.length > 0) {
    allEmails = allEmails.concat(aliases);
  }
  return allEmails;
}

// 5. Main Function: Process Emails
function processEmails(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Recipients");

  if (!sheet) return "Error: 'Recipients' sheet not found.";

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const primaryEmail = Session.getActiveUser().getEmail();
  
  // Use the selected template, or default to "templatesatu" if something goes wrong
  const selectedTemplate = formData.templateSelect || "templatesatu";
  
  let sentCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0];
    const email = row[1];
    const status = row[2];

    if (email && status !== "Sent") {
      const statusCell = sheet.getRange(i + 1, 3);
      
      try {
        // --- DYNAMIC TEMPLATE SELECTION ---
        const htmlTemplate = HtmlService.createTemplateFromFile(selectedTemplate);
        htmlTemplate.name = name; 
        const htmlBody = htmlTemplate.evaluate().getContent();

        let emailOptions = {
          name: formData.senderName,
          replyTo: formData.replyTo,
          htmlBody: htmlBody
        };

        // Handle "From" Alias safely
        if (formData.fromEmail && formData.fromEmail !== "" && formData.fromEmail !== primaryEmail) {
          emailOptions.from = formData.fromEmail;
        }

        GmailApp.sendEmail(email, formData.subject, "HTML not supported", emailOptions);

        statusCell.setValue("Sent").setBackground("#d9ead3");
        sentCount++;
        Utilities.sleep(500);

      } catch (e) {
        statusCell.setValue("Error: " + e.message).setBackground("#f4cccc");
        console.error("Error sending to " + email, e);
      }
    }
  }
  return "Finished! Sent " + sentCount + " emails using " + selectedTemplate;
}
