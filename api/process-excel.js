const XLSX = require('xlsx');

exports.handler = async (event, context) => {
  // Enable CORS
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  // Handle preflight requests
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers,
      body: ''
    };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    // Parse the multipart form data
    const body = event.body;
    const boundary = event.headers['content-type'].split('boundary=')[1];
    
    if (!boundary || !body) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'No file uploaded' })
      };
    }

    // Extract file data from multipart body
    const fileData = extractFileFromMultipart(body, boundary);
    
    if (!fileData) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'Could not extract file data' })
      };
    }

    // Parse Excel file
    const workbook = XLSX.read(fileData, { type: 'buffer' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Filter for laptops/MacBooks
    const laptops = jsonData.filter(item => 
      item['Sub-Category'] && 
      item['Sub-Category'].toLowerCase().includes('laptop')
    );

    // IMPROVED PRODUCT GROUPING LOGIC
    const productGroups = groupProductsImproved(laptops);

    const processedData = {
      totalItems: laptops.length,
      productGroups: productGroups,
      rawData: laptops,
      groupCount: Object.keys(productGroups).length
    };

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(processedData)
    };

  } catch (error) {
    console.error('Error processing Excel:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: 'Error processing Excel file: ' + error.message })
    };
  }
};

// IMPROVED GROUPING LOGIC - Creates fewer, more logical product groups
function groupProductsImproved(laptops) {
  const productGroups = {};

  laptops.forEach(item => {
    // Create a more specific product key that groups similar configurations
    const productKey = createImprovedProductKey(item);

    if (!productGroups[productKey]) {
      productGroups[productKey] = {
        model: normalizeModel(item.Model),
        processor: normalizeProcessor(item.Processor),
        storage: normalizeStorage(item.Storage),
        memory: normalizeMemory(item.Memory),
        basePrice: extractPrice(item),
        items: [],
        variants: {}
      };
    }

    // Group by condition and color for variants
    const variantKey = `${item.Color || 'Default'}_${item.Condition || 'Unknown'}`;
    
    if (!productGroups[productKey].variants[variantKey]) {
      productGroups[productKey].variants[variantKey] = {
        color: item.Color || 'Default',
        condition: item.Condition || 'Unknown',
        quantity: 0,
        items: []
      };
    }

    productGroups[productKey].variants[variantKey].quantity++;
    productGroups[productKey].variants[variantKey].items.push(item);
    productGroups[productKey].items.push(item);
  });

  return productGroups;
}

function createImprovedProductKey(item) {
  const model = normalizeModel(item.Model || 'Unknown');
  const processor = normalizeProcessor(item.Processor || '');
  const storage = normalizeStorage(item.Storage || '');
  const memory = normalizeMemory(item.Memory || '');

  // Create a more logical grouping key
  return `${model}_${processor}_${storage}_${memory}`;
}

function normalizeModel(model) {
  if (!model) return 'Unknown';
  
  // Extract main model info
  const modelStr = model.toString().toLowerCase();
  
  if (modelStr.includes('macbook pro')) return 'MacBook Pro';
  if (modelStr.includes('macbook air')) return 'MacBook Air';
  if (modelStr.includes('macbook')) return 'MacBook';
  if (modelStr.includes('imac')) return 'iMac';
  if (modelStr.includes('mac mini')) return 'Mac Mini';
  
  return model;
}

function normalizeProcessor(processor) {
  if (!processor) return 'Unknown';
  
  const procStr = processor.toString().toLowerCase();
  
  if (procStr.includes('m3')) return 'M3';
  if (procStr.includes('m2')) return 'M2';
  if (procStr.includes('m1')) return 'M1';
  if (procStr.includes('intel') || procStr.includes('i7')) return 'Intel i7';
  if (procStr.includes('i5')) return 'Intel i5';
  if (procStr.includes('i3')) return 'Intel i3';
  
  return processor.substring(0, 10);
}

function normalizeStorage(storage) {
  if (!storage) return 'Unknown';
  
  const storageStr = storage.toString().toLowerCase();
  
  if (storageStr.includes('1tb') || storageStr.includes('1000gb')) return '1TB';
  if (storageStr.includes('512gb')) return '512GB';
  if (storageStr.includes('256gb')) return '256GB';
  if (storageStr.includes('128gb')) return '128GB';
  if (storageStr.includes('2tb')) return '2TB';
  
  return storage;
}

function normalizeMemory(memory) {
  if (!memory) return 'Unknown';
  
  const memStr = memory.toString().toLowerCase();
  
  if (memStr.includes('32gb')) return '32GB';
  if (memStr.includes('16gb')) return '16GB';
  if (memStr.includes('8gb')) return '8GB';
  if (memStr.includes('4gb')) return '4GB';
  if (memStr.includes('64gb')) return '64GB';
  
  return memory;
}

function extractPrice(item) {
  // Try to extract price from various possible fields
  const priceFields = ['Price', 'Cost', 'Value', 'Amount'];
  
  for (const field of priceFields) {
    if (item[field] && !isNaN(parseFloat(item[field]))) {
      return parseFloat(item[field]);
    }
  }
  
  return 0;
}

function extractFileFromMultipart(body, boundary) {
  try {
    const buffer = Buffer.from(body, 'base64');
    const boundaryBuffer = Buffer.from(`--${boundary}`);
    
    // Find file content between boundaries
    const parts = buffer.toString('binary').split(`--${boundary}`);
    
    for (const part of parts) {
      if (part.includes('filename=') && part.includes('Content-Type:')) {
        // Extract the binary data
        const headerEnd = part.indexOf('\r\n\r\n');
        if (headerEnd !== -1) {
          const fileContent = part.substring(headerEnd + 4);
          return Buffer.from(fileContent, 'binary');
        }
      }
    }
    
    return null;
  } catch (error) {
    console.error('Error extracting file:', error);
    return null;
  }
}