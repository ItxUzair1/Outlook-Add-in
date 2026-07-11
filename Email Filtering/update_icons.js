const fs = require('fs');

const manifests = [
  'manifest.addin.xml',
  'manifest_local.xml',
  'manifest_test.xml',
  'manifest-client-test.xml'
];

const mappings = {
  'Email': 'email',
  'Multi': 'multi',
  'Location': 'location',
  'Search': 'search',
  'Tools': 'option',
  'Help': 'help'
};

for (const file of manifests) {
  if (!fs.existsSync(file)) continue;
  let content = fs.readFileSync(file, 'utf8');
  
  for (const [key, iconName] of Object.entries(mappings)) {
    for (const size of ['16', '32', '80']) {
      const regex = new RegExp(`(id="Icon\.${key}\.${size}" DefaultValue="[^"]*/assets/)[^"]+\\.png"`, 'g');
      content = content.replace(regex, `$1${iconName}-${size}.png"`);
    }
  }
  
  fs.writeFileSync(file, content);
  console.log(`Updated XML: ${file}`);
}

// Now update manifest.json
const jsonFile = 'manifest.json';
if (fs.existsSync(jsonFile)) {
  const data = JSON.parse(fs.readFileSync(jsonFile, 'utf8'));
  
  function traverse(obj) {
    if (Array.isArray(obj)) {
      obj.forEach(traverse);
    } else if (typeof obj === 'object' && obj !== null) {
      if (obj.icons && Array.isArray(obj.icons)) {
        // Find which feature this is
        // Usually sibling is "id"
        let feature = obj.id || '';
        let iconName = null;
        if (feature.includes('FileEmail')) iconName = 'email';
        else if (feature.includes('FileMultiple')) iconName = 'multi';
        else if (feature.includes('Search')) iconName = 'search';
        else if (feature.includes('Collections')) iconName = 'location';
        else if (feature.includes('Options')) iconName = 'option';
        else if (feature.includes('Help')) iconName = 'help';
        
        if (iconName) {
          obj.icons.forEach(icon => {
            if ([16, 32, 80].includes(icon.size)) {
              icon.url = icon.url.replace(/\/assets\/[^/]+$/, `/assets/${iconName}-${icon.size}.png`);
            }
          });
        }
      }
      Object.values(obj).forEach(traverse);
    }
  }
  
  traverse(data);
  fs.writeFileSync(jsonFile, JSON.stringify(data, null, 4));
  console.log(`Updated JSON: ${jsonFile}`);
}
