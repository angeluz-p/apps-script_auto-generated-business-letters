// Add these caching functions
const employeeCache = new Map();
const templateCache = new Map();
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

function getAllEmployeesCached() {
  const now = Date.now();
  const cached = employeeCache.get('employees');
  
  if (cached && (now - cached.timestamp) < CACHE_DURATION) {
    return cached.data;
  }
  
  const employees = getAllEmployees();
  employeeCache.set('employees', { data: employees, timestamp: now });
  return employees;
}

function getAllTemplateNamesCached() {
  const now = Date.now();
  const cached = templateCache.get('templates');
  
  if (cached && (now - cached.timestamp) < CACHE_DURATION) {
    return cached.data;
  }
  
  const templates = getAllTemplateNames();
  templateCache.set('templates', { data: templates, timestamp: now });
  return templates;
}

function warmCaches() {
  getAllEmployeesCached();
  getAllTemplateNamesCached();
}