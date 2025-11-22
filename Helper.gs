// Function to convert HTML to plain text for Google Docs
function convertHtmlToPlainText(html) {
  if (!html) return '';
  
  // Convert HTML formatting to plain text equivalents
  let text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<p[^>]*>/gi, '')
    .replace(/<\/div>/gi, '\n')
    .replace(/<div[^>]*>/gi, '')
    .replace(/<strong[^>]*>(.*?)<\/strong>/gi, '$1')  // Remove bold formatting but keep text
    .replace(/<b[^>]*>(.*?)<\/b>/gi, '$1')           // Remove bold formatting but keep text
    .replace(/<em[^>]*>(.*?)<\/em>/gi, '$1')         // Remove italic formatting but keep text
    .replace(/<i[^>]*>(.*?)<\/i>/gi, '$1')           // Remove italic formatting but keep text
    .replace(/<ul[^>]*>/gi, '')
    .replace(/<\/ul>/gi, '\n')
    .replace(/<ol[^>]*>/gi, '')
    .replace(/<\/ol>/gi, '\n')
    .replace(/<li[^>]*>/gi, '       ●    ')          // Convert list items to bullet points
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]*>/g, '')                         // Remove any remaining HTML tags
    .replace(/&nbsp;/gi, ' ')                        // Convert non-breaking spaces
    .replace(/&amp;/gi, '&')                         // Convert HTML entities
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\n\s*\n/g, '\n\n')                     // Clean up multiple line breaks
    .trim();
  
  return text;
}

// Function to find the column number for "Letter Type"
function findLetterTypeColumn() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim().toLowerCase() === "letter type") {
        return i + 1;
      }
    }
    return null;
  } catch (error) {
    console.error("Error finding Letter Type column:", error);
    return null;
  }
}

// Function to format currency values
function formatCurrency(value) {
  if (!value || value === "" || isNaN(value)) {
    return "";
  }
  const numValue = parseFloat(value);
  return numValue.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

// Function to determine if a field is a date field
function isDateField(fieldName) {
  const dateKeywords = ['date', 'Date', 'DATE', 'effective', 'Effective', 'start', 'Start', 'end', 'End', 'violation', 'Violation', 'deadline', 'Deadline', 'submission', 'Submission'];
  return dateKeywords.some(keyword => fieldName.includes(keyword));
}

/** 
 * DATE FORMATS
*/

// Function for numbers ordinal (e.g., st, nd)
function getOrdinalSuffix(num) {
  const j = num % 10;
  const k = num % 100;
  
  if (j == 1 && k != 11) {
    return num + "st";
  }
  if (j == 2 && k != 12) {
    return num + "nd";
  }
  if (j == 3 && k != 13) {
    return num + "rd";
  }
  return num + "th";
}

// Function to get multiple date formats for a single date
function getMultipleDateFormats(dateObj, baseName) {
  if (!dateObj || !(dateObj instanceof Date)) {
    return {};
  }

  const timezone = Session.getScriptTimeZone();
  const formats = {};
  
  // Standard formats
  formats[baseName] = Utilities.formatDate(dateObj, timezone, "MMMM dd, yyyy");
  formats[baseName + "Short"] = Utilities.formatDate(dateObj, timezone, "MM/dd/yyyy");
  formats[baseName + "Long"] = Utilities.formatDate(dateObj, timezone, "EEEE, MMMM dd, yyyy");
  formats[baseName + "MonthYear"] = Utilities.formatDate(dateObj, timezone, "MMMM yyyy");
  formats[baseName + "Year"] = Utilities.formatDate(dateObj, timezone, "yyyy");
  formats[baseName + "Month"] = Utilities.formatDate(dateObj, timezone, "MMMM");
  formats[baseName + "Day"] = Utilities.formatDate(dateObj, timezone, "dd");
  formats[baseName + "ISO"] = Utilities.formatDate(dateObj, timezone, "yyyy-MM-dd");
  formats[baseName + "Numeric"] = Utilities.formatDate(dateObj, timezone, "MM/dd/yyyy");
  
  // Ordinal day formats
  const dayNumber = parseInt(Utilities.formatDate(dateObj, timezone, "dd"));
  const ordinalDay = getOrdinalSuffix(dayNumber);
  formats[baseName + "DayOrdinal"] = ordinalDay;
  
  return formats;
}

// Function to format dates based on field type or name
function formatDateByType(dateObj, fieldName) {
  if (!dateObj || !(dateObj instanceof Date)) {
    return "";
  }

  const timezone = Session.getScriptTimeZone();
  
  // Different date formats based on field name or type
  switch (true) {
    case fieldName.toLowerCase().includes('start') || fieldName.toLowerCase().includes('end'):
      // Employment dates - use month/year format
      return Utilities.formatDate(dateObj, timezone, "MMMM yyyy");
      
    case fieldName.toLowerCase().includes('effective'):
      // Effective dates - use full date
      return Utilities.formatDate(dateObj, timezone, "MMMM dd, yyyy");
      
    case fieldName.toLowerCase().includes('violation'):
      // Violation dates - use short format
      return Utilities.formatDate(dateObj, timezone, "MM/dd/yyyy");
      
    case fieldName.toLowerCase().includes('birth'):
      // Birth dates - use full date
      return Utilities.formatDate(dateObj, timezone, "MMMM dd, yyyy");
      
    default:
      // Default format
      return Utilities.formatDate(dateObj, timezone, "MMMM dd, yyyy");
  }
}

/** 
 * TITLE CASE FORMATTING FUNCTIONS
*/

// Function to convert text to Title Case
function toTitleCase(str) {
  if (!str || typeof str !== 'string') return str;
  
  return str.toLowerCase().replace(/\b\w+/g, function(word) {
    // Handle special cases for names
    const specialCases = {
      'ii': 'II',
      'iii': 'III',
      'iv': 'IV',
      'jr': 'Jr.',
      'sr': 'Sr.',
      'van': 'van',
      'von': 'von',
      'de': 'de',
      'del': 'del',
      'la': 'la',
      'le': 'le',
      'da': 'da',
      'di': 'di',
      'du': 'du',
      'mc': 'Mc',
      'mac': 'Mac',
      'o\'': 'O\'',
      'd\'': 'D\''
    };
    
    const lowerWord = word.toLowerCase();
    
    // Check for special cases first
    if (specialCases[lowerWord]) {
      return specialCases[lowerWord];
    }
    
    // Handle McDonald, MacArthur type names
    if (lowerWord.startsWith('mc') && lowerWord.length > 2) {
      return 'Mc' + lowerWord.charAt(2).toUpperCase() + lowerWord.slice(3);
    }
    
    if (lowerWord.startsWith('mac') && lowerWord.length > 3) {
      return 'Mac' + lowerWord.charAt(3).toUpperCase() + lowerWord.slice(4);
    }
    
    // Handle O'Connor, D'Angelo type names
    if (lowerWord.includes('\'')) {
      return lowerWord.split('\'').map(part => 
        part.charAt(0).toUpperCase() + part.slice(1)
      ).join('\'');
    }
    
    // Handle hyphenated names
    if (lowerWord.includes('-')) {
      return lowerWord.split('-').map(part => 
        part.charAt(0).toUpperCase() + part.slice(1)
      ).join('-');
    }
    
    // Default title case
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  });
}

// Function to determine if a field contains names
function isNameField(fieldName) {
  // First check if it's a company-related field that should be excluded
  const companyKeywords = [
    'company', 'Company', 'COMPANY',
    'companyname', 'CompanyName', 'COMPANYNAME',
    'organization', 'Organization', 'ORGANIZATION',
    'ltd', 'Ltd', 'LTD',
    'llc', 'LLC'
  ];
  
  // If it's a company field, don't apply name formatting
  if (companyKeywords.some(keyword => 
    fieldName.toLowerCase().includes(keyword.toLowerCase())
  )) {
    return false;
  }
  
  // Check for person name fields
  const nameKeywords = [
    'name', 'Name', 'NAME',
    'first', 'First', 'FIRST',
    'last', 'Last', 'LAST',
    'full', 'Full', 'FULL',
    'employee', 'Employee', 'EMPLOYEE',
    'recipient', 'Recipient', 'RECIPIENT',
    'approver', 'Approver', 'APPROVER',
    'manager', 'Manager', 'MANAGER',
    'supervisor', 'Supervisor', 'SUPERVISOR',
    'hr', 'HR', 'human resource', 'Human Resource',
    'reported by', 'Reported By', 'REPORTED BY',
    'issued by', 'Issued By', 'ISSUED BY',
    'turnover to', 'Turnover To', 'TURNOVER TO'
  ];
  
  return nameKeywords.some(keyword => 
    fieldName.toLowerCase().includes(keyword.toLowerCase())
  );
}

// Function to convert job titles/positions to proper case
function toPositionTitleCase(str) {
  if (!str || typeof str !== 'string') return str;
  
  // Words that should remain lowercase in titles (unless at beginning)
  const lowercaseWords = [
    'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'if', 'in', 'into',
    'is', 'it', 'no', 'not', 'of', 'on', 'or', 'such', 'that', 'the',
    'their', 'then', 'there', 'these', 'they', 'this', 'to', 'was', 'will', 'with'
  ];
  
  // Common abbreviations that should be uppercase
  const uppercaseAbbreviations = {
    'ceo': 'CEO',
    'cfo': 'CFO',
    'cto': 'CTO',
    'coo': 'COO',
    'hr': 'HR',
    'it': 'IT',
    'qa': 'QA',
    'ui': 'UI',
    'ux': 'UX',
    'api': 'API',
    'seo': 'SEO',
    'ppc': 'PPC',
    'roi': 'ROI',
    'kpi': 'KPI',
    'crm': 'CRM',
    'erp': 'ERP',
    'sme': 'SME',
    'vp': 'VP',
    'svp': 'SVP',
    'evp': 'EVP',
    'avp': 'AVP',
    'gm': 'GM',
    'am': 'AM',
    'pm': 'PM',
    'ba': 'BA',
    'sa': 'SA',
    'dba': 'DBA',
    'sqa': 'SQA',
    'ii': 'II',
    'iii': 'III',
    'iv': 'IV',
    'jr': 'Jr.',
    'sr': 'Sr.'
  };
  
  return str.toLowerCase().replace(/\b\w+/g, function(word, index, fullStr) {
    const lowerWord = word.toLowerCase();
    const isFirstWord = index === 0;
    const isLastWord = index + word.length === fullStr.length;
    
    // Check for abbreviations first
    if (uppercaseAbbreviations[lowerWord]) {
      return uppercaseAbbreviations[lowerWord];
    }
    
    // Always capitalize first and last words
    if (isFirstWord || isLastWord) {
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }
    
    // Keep certain words lowercase unless they're first/last
    if (lowercaseWords.includes(lowerWord)) {
      return lowerWord;
    }
    
    // Default title case
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  });
}

// Function to determine if a field contains job positions/titles
function isPositionField(fieldName) {
  const positionKeywords = [
    'position', 'Position', 'POSITION',
    'title', 'Title', 'TITLE',
    'job title', 'Job Title', 'JOB TITLE',
    'role', 'Role', 'ROLE',
    'designation', 'Designation', 'DESIGNATION',
    'rank', 'Rank', 'RANK',
    'report to', 'Report To', 'REPORT TO'
  ];
  
  return positionKeywords.some(keyword => 
    fieldName.toLowerCase().includes(keyword.toLowerCase())
  );
}

// Function to convert locations to proper case
function toLocationTitleCase(str) {
  if (!str || typeof str !== 'string') return str;
  
  // Words that should remain lowercase in locations (unless at beginning)
  const lowercaseWords = [
    'of', 'the', 'and', 'at', 'in', 'on', 'by', 'to', 'from', 'with', 'de', 'del', 'la', 'le', 'da', 'di'
  ];
  
  // Common location abbreviations and special cases
  const locationSpecialCases = {
    'usa': 'USA',
    'uk': 'UK',
    'uae': 'UAE',
    'ph': 'PH',
    'hq': 'HQ',
    'bldg': 'Bldg.',
    'blvd': 'Blvd.',
    'ave': 'Ave.',
    'st': 'St.',
    'rd': 'Rd.',
    'dr': 'Dr.',
    'ln': 'Ln.',
    'ct': 'Ct.',
    'pl': 'Pl.',
    'sq': 'Sq.',
    'cir': 'Cir.',
    'pkwy': 'Pkwy.',
    'n': 'N.',
    'e': 'E.',
    's': 'S.',
    'w': 'W.',
    'ne': 'NE',
    'nw': 'NW',
    'se': 'SE',
    'sw': 'SW',
    'floor': 'Floor',
    'flr': 'Flr.',
    'suite': 'Suite',
    'ste': 'Ste.',
    'unit': 'Unit',
    'apt': 'Apt.',
    'rm': 'Rm.',
    'room': 'Room'
  };
  
  // Handle numbered streets (1st, 2nd, 3rd, etc.)
  const ordinalRegex = /(\d+)(st|nd|rd|th)\b/gi;
  
  return str.toLowerCase().replace(/\b\w+/g, function(word, index, fullStr) {
    const lowerWord = word.toLowerCase();
    const isFirstWord = index === 0;
    const isLastWord = index + word.length === fullStr.length;
    
    // Check for special location cases first
    if (locationSpecialCases[lowerWord]) {
      return locationSpecialCases[lowerWord];
    }
    
    // Handle ordinal numbers (1st, 2nd, 3rd, etc.)
    if (ordinalRegex.test(word)) {
      return word.replace(ordinalRegex, function(match, num, suffix) {
        return num + suffix.toLowerCase();
      });
    }
    
    // Always capitalize first and last words
    if (isFirstWord || isLastWord) {
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }
    
    // Keep certain words lowercase unless they're first/last
    if (lowercaseWords.includes(lowerWord)) {
      return lowerWord;
    }
    
    // Default title case
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  });
}

// Function to determine if a field contains location information
function isLocationField(fieldName) {
  const locationKeywords = [
    'location', 'Location', 'LOCATION',
    'address', 'Address', 'ADDRESS',
    'city', 'City', 'CITY',
    'province', 'Province', 'PROVINCE',
    'state', 'State', 'STATE',
    'country', 'Country', 'COUNTRY',
    'office', 'Office', 'OFFICE',
    'site', 'Site', 'SITE',
    'branch', 'Branch', 'BRANCH',
    'region', 'Region', 'REGION',
    'company address', 'Company Address', 'COMPANY ADDRESS',
    'work location', 'Work Location', 'WORK LOCATION',
    'incident location', 'Incident Location', 'INCIDENT LOCATION'
  ];
  
  return locationKeywords.some(keyword => 
    fieldName.toLowerCase().includes(keyword.toLowerCase())
  );
}