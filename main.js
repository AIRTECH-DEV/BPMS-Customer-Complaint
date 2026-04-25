function handleFormSubmit(e) {
  const sheetName = e.range.getSheet().getName();

  // 1️⃣ Assign complaint ID (only affects Complaint Entry sheet)
  assignComplaintId(e);

  // 2️⃣ Route based on which sheet was submitted
  const HANDLERS = {
    "Complaint Entry": handleComplaintForm, // complaint form
    //"Complaint Report": handleComplaintReportForm,// report upload form
    "Complaint Report1" :submitComplaint
    
  };

  const handler = HANDLERS[sheetName];
  if (handler) {
    handler(e);
  }

  // 3️⃣ GLOBAL POST-PROCESSING (ONLY after Complaint Report submission)
  if (sheetName.trim().toLowerCase() === "complaint report") {
    try {
      generatePaymentFollowup();
      generateWarrantyPending();
      generateLabourAMC();
      processFeedbackFromReport(e)
      Logger.log("Post-processing executed for Complaint Report.");
    } catch (err) {
      Logger.log("Error in post-processing: " + err);
    }
  }
}
