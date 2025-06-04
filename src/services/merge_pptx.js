#!/usr/bin/env node
// This file is compiled into the dist/bundle.js file
// It is used to merge the pptx files in the folder into a single pptx file
// The output file is named fusion.pptx
// The input folder is the folder containing the pptx files to merge
// The output folder is the folder where the merged pptx file will be saved

const Automizer = require('pptx-automizer').default;
const fs        = require('fs');
const path      = require('path');

async function MergePptx(folderPath, outputPath) {
  // 1. Create output folder if necessary
  if (!fs.existsSync(outputPath)) {
    fs.mkdirSync(outputPath, { recursive: true });
  }
  
  // 2. Read and sort .pptx files
  const files = fs.readdirSync(folderPath)
    .filter(f => f.toLowerCase().endsWith('.pptx'))
    .sort();

  if (files.length === 0) {
    throw new Error(`No .pptx files found in ${folderPath}`);
  }

  // 3. First file as root
  const rootFile = files.shift();
  
  const pres = new Automizer({
    templateDir: folderPath,
    outputDir: outputPath
  })
  .loadRoot(path.join(folderPath, rootFile));

  // 4. Loop through the rest of the files
  for (const file of files) {
    const key = path.basename(file, '.pptx');
    pres.load(path.join(folderPath, file), key);

    // Get and add all slide numbers
    const slideNumbers = await pres
      .getTemplate(key)
      .getAllSlideNumbers();

    for (const num of slideNumbers) {
      pres.addSlide(key, num);
    }
  }

  // 5. Write the result
  await pres.write(path.join( 'fusion.pptx'));
  console.log('Merge completed in', path.join( 'fusion.pptx'));
}

function printUsage() {
  console.log(`
Usage:
  node merge_pptx.js <dossier_source> <dossier_sortie>

Exemple:
  node merge_pptx.js ./pptx_folder ./test
`);
}

// Get args
const [,, folderPath, outputPath] = process.argv;

if (!folderPath || !outputPath) {
  printUsage();
  process.exit(1);
}

// Launch
MergePptx(folderPath, outputPath)
  .catch(err => {
    console.error('Error during fusion :', err.message);
    process.exit(1);
  });
