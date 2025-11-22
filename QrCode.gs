// Generate QR code URL using QuickChart.io
function generateQRCode(data) {
  try {
    const baseUrl = "https://quickchart.io/qr";
    const params = `?text=${encodeURIComponent(data)}&size=150&format=png`;
    const fullUrl = baseUrl + params;
    
    console.log("Generated QR URL:", fullUrl);
    
    // Test the URL by fetching it
    const response = UrlFetchApp.fetch(fullUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(`QR service returned ${response.getResponseCode()}`);
    }
    
    return fullUrl;
  } catch (error) {
    console.error("QR generation error:", error);
    // Fallback to a different service
    return `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(data)}`;
  }
}

// Generate serial number for certificates
function generateSerialNumber(letterType) {
  const year = new Date().getFullYear();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // You can create a tracking sheet or use existing data
  // For now, using timestamp for uniqueness
  const timestamp = new Date().getTime().toString().slice(-6);
  
  const prefixes = {
    'COE': 'COE',
    'Certificate of Employment': 'COE'
  };
  
  const prefix = prefixes[letterType] || 'DOC';
  return `${prefix}-${year}-${timestamp}`;
}

// Create verification URL for QR code
function createVerificationURL(serialNumber) {
  // Replace with your actual web app URL when deployed
  const webAppUrl = "https://script.google.com/macros/s/{ID}/exec";
  return `${webAppUrl}?verify=${serialNumber}`;
}