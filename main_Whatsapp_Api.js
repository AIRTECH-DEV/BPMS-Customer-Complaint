function onFormSubmitMaster(e) {
  const sheetName = e.range.getSheet().getName();
  
  if (sheetName === "Complaint Entry") { // Your first form sheet
    Com_register(e); 
 
}
}
