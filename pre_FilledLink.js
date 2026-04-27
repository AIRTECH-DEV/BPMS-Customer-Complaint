

function onFormSubmit(e) {
  const responses = e.namedValues;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  
  const getVal = (headerName) => {
    return (responses[headerName] && responses[headerName][0]) ? responses[headerName][0].trim() : "";
  };

  // --- 1. COLLECT EMAILS ---
  let emailList = [
    getVal('Email Address'),
    getVal('Email ID'),
    getVal('Contact Person Email ID')
  ];
  let uniqueEmails = [...new Set(emailList.filter(email => email && email.includes('@')))];

  const baseUrl = "https://docs.google.com/forms/d/e/1FAIpQLScqVA0uc2ka91Vjc0voOIhRWhrH8qbMPXLn5GwoGzId8f0j2w/viewform?usp=pp_url";
  
  // --- 2. URL MAPPING ---
  const customerType = getVal('Customer Type');
  
  // Start with common fields
  let prefilledUrl = baseUrl + 
    "&entry.1063906463=" + encodeURIComponent(getVal('Service Type')) +
    "&entry.1863661142=" + encodeURIComponent(customerType);

  if (customerType.includes("Commercial")) {
    // COMMERCIAL SPECIFIC IDs
    prefilledUrl += 
      "&entry.364759650="  + encodeURIComponent(getVal('Company Name')) +
      "&entry.1743700807=" + encodeURIComponent(getVal('Site Flat / Shop / Showroom No')) +
      "&entry.2016624238=" + encodeURIComponent(getVal('Site Building Name')) +
      "&entry.1794926247=" + encodeURIComponent(getVal('Street / Locality / Area')) + 
      "&entry.1854736678=" + encodeURIComponent(getVal('City')) + 
      "&entry.1463003551=" + encodeURIComponent(getVal('Pincode')) + 
      "&entry.155186065="  + encodeURIComponent(getVal('Complaint Type')) + 
      "&entry.1252578666=" + encodeURIComponent(getVal('Machine Brand')) +
      "&entry.943254356="  + encodeURIComponent(getVal('Machine / System Type')) +
      "&entry.1691676309=" + encodeURIComponent(getVal('Contact Person Name')) +
      "&entry.50000003="   + encodeURIComponent(getVal('Contact Person phone Number') || getVal('Contact Person Number')) +
      "&entry.1198011696=" + encodeURIComponent(getVal('Contact Person Email ID'));
  } else {
    // RESIDENTIAL SPECIFIC IDs (Updated based on your Residential link)
    prefilledUrl += 
      "&entry.1309771941=" + encodeURIComponent(getVal('Client Name')) +
      "&entry.1084613101=" + encodeURIComponent(getVal('Flat No')) + 
      "&entry.2052584446=" + encodeURIComponent(getVal('Building')) + 
      "&entry.766022231="  + encodeURIComponent(getVal('Street / Locality / Area')) + 
      "&entry.1417254603=" + encodeURIComponent(getVal('City')) + 
      "&entry.692806498="  + encodeURIComponent(getVal('Pincode')) + 
      "&entry.1974749771=" + encodeURIComponent(getVal('Complaint Type')) +
      "&entry.1254142612=" + encodeURIComponent(getVal('Machine Brand')) +
      "&entry.853699598="  + encodeURIComponent(getVal('Machine / System Type')) +
      "&entry.1828559079=" + encodeURIComponent(getVal('Contact Person Name')) +
      "&entry.1688807771=" + encodeURIComponent(getVal('Contact Person phone Number') || getVal('Contact Person Number')) +
      "&entry.985548176="  + encodeURIComponent(getVal('Contact Person Email ID'));
  }

  // --- 3. SEND EMAIL ---
  if (uniqueEmails.length > 0) {
    const siteName = getVal('Company Name') || getVal('Client Name') || "Valued Customer";
    const subject = "Service Link for: " + siteName;
    const htmlBody = `<p>Dear Customer,</p><p>Please use the link below for future service requests: </p><p><strong><a href="${prefilledUrl}">Click Here</a></strong></p><p>Regards,<br>Vakharia Airtech</p>`;

    uniqueEmails.forEach(email => {
      try {
        GmailApp.sendEmail(email, subject, "", {
          from: 'customercare@vakhariaairtech.com',
          name: 'Vakharia Airtech Customer Care',
          htmlBody: htmlBody
        });
      } catch (err) { console.error("Email error: " + err); }
    });
  }

  // --- 4. SAVE TO SHEET ---
  sheet.getRange(row, 32).setValue(prefilledUrl);
}