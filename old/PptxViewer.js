// module that utilise pptx-in-html-out to convert pptx to html and then use reveal.js to display pptx files inside a browser.

// PPTX web editor using PptxGenJS and js-pptx for editing PowerPoint files in browser
import React, { useState, useCallback, useEffect, useRef } from 'react';
import { useDropzone } from 'react-dropzone';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import './PptxViewer.css';

const PptxViewer = () => {
  const [slides, setSlides] = useState([]);
  const [selectedSlide, setSelectedSlide] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [fileName, setFileName] = useState('');
  const previewRefs = useRef([]);

  // Handle file drop
  const onDrop = useCallback(async (acceptedFiles) => {
    const file = acceptedFiles[0];
    if (!file) return;
    
    setIsLoading(true);
    setFileName(file.name);
    
    try {
      const content = await readFileAsArrayBuffer(file);
      const parsedSlides = await parsePptx(content);
      setSlides(parsedSlides);
      setSelectedSlide(0);
    } catch (error) {
      console.error('Error processing PPTX file:', error);
      alert('Failed to process the PPTX file');
    } finally {
      setIsLoading(false);
    }
  }, []);
  
  const { getRootProps, getInputProps, isDragActive } = useDropzone({ 
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx']
    }
  });

  // Update preview refs when slides change
  useEffect(() => {
    previewRefs.current = previewRefs.current.slice(0, slides.length);
  }, [slides]);

  // Read file as ArrayBuffer
  const readFileAsArrayBuffer = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(e);
      reader.readAsArrayBuffer(file);
    });
  };

  // Parse PPTX file and extract slides with improved structure preservation
  const parsePptx = async (content) => {
    try {
      const zip = new JSZip();
      const zipData = await zip.loadAsync(content);
      
      // Get presentation.xml to extract master elements and theme information
      let presentationXml = '';
      if (zipData.files['ppt/presentation.xml']) {
        presentationXml = await zipData.files['ppt/presentation.xml'].async('text');
      }
      
      // Extract relationships to find master slide references
      let relationshipsXml = '';
      if (zipData.files['ppt/_rels/presentation.xml.rels']) {
        relationshipsXml = await zipData.files['ppt/_rels/presentation.xml.rels'].async('text');
      }
      
      // Get slide layout information
      const slideLayouts = {};
      const slideLayoutFiles = Object.keys(zipData.files).filter(fname => fname.match(/ppt\/slideLayouts\/slideLayout[0-9]+\.xml/));
      for (const layoutFile of slideLayoutFiles) {
        try {
          const layoutContent = await zipData.files[layoutFile].async('text');
          const layoutIdMatch = layoutFile.match(/slideLayout([0-9]+)\.xml/);
          if (layoutIdMatch) {
            const layoutId = layoutIdMatch[1];
            slideLayouts[layoutId] = layoutContent;
          }
        } catch (e) {
          console.error(`Error reading layout ${layoutFile}:`, e);
        }
      }
      
      // Extract slide data from the PPTX XML structure with improved ordering
      const slideEntries = Object.keys(zipData.files)
        .filter(filename => filename.match(/ppt\/slides\/slide[0-9]+\.xml/))
        .sort((a, b) => {
          const numA = parseInt(a.match(/slide([0-9]+)\.xml/)[1]);
          const numB = parseInt(b.match(/slide([0-9]+)\.xml/)[1]);
          return numA - numB;
        });
      
      // Process each slide
      const parsedSlides = [];
      
      for (const slideEntry of slideEntries) {
        const slideContent = await zipData.files[slideEntry].async('text');
        
        // Extract slide ID to find related media
        const slideIdMatch = slideEntry.match(/slide([0-9]+)\.xml/);
        const slideId = slideIdMatch ? slideIdMatch[1] : '';
        
        // Check for relationship file for this slide
        let slideRelsXml = '';
        const slideRelsPath = `ppt/slides/_rels/slide${slideId}.xml.rels`;
        if (zipData.files[slideRelsPath]) {
          slideRelsXml = await zipData.files[slideRelsPath].async('text');
        }
        
        // Determine slide layout reference if available
        let layoutRef = '';
        const layoutRefMatch = slideContent.match(/layoutId="([^"]+)"/);
        if (layoutRefMatch) {
          layoutRef = layoutRefMatch[1];
        }
        
        // Extract title with priority order:
        // 1. Using cSld name attribute
        // 2. Using title placeholder content
        // 3. Using first text box if nothing else found
        let title = '';
        
        const titleFromCsldMatch = slideContent.match(/<p:cSld[^>]*name="([^"]+)"/);
        if (titleFromCsldMatch && titleFromCsldMatch[1]) {
          title = titleFromCsldMatch[1];
        }
        
        if (!title) {
          // Try to extract from title placeholder - these have type="title" in their placeholders
          const titlePlaceholderMatch = slideContent.match(/<p:ph[^>]*type="title"[^>]*>[\s\S]*?<a:t>(.*?)<\/a:t>/);
          if (titlePlaceholderMatch && titlePlaceholderMatch[1]) {
            title = titlePlaceholderMatch[1];
          } else {
            // Try the idx=0 placeholder which is often the title
            const titleIdxMatch = slideContent.match(/<p:ph[^>]*idx="0"[^>]*>[\s\S]*?<a:t>(.*?)<\/a:t>/);
            if (titleIdxMatch && titleIdxMatch[1]) {
              title = titleIdxMatch[1];
            }
          }
        }
        
        // If still no title, look for the first substantial text element
        if (!title) {
          const firstTextMatch = slideContent.match(/<a:t>(.*?)<\/a:t>/);
          if (firstTextMatch && firstTextMatch[1] && firstTextMatch[1].trim().length > 0) {
            title = firstTextMatch[1];
          } else {
            title = `Slide ${slideId}`;
          }
        }
        
        // Extract structured content to maintain hierarchy
        const structuredContent = [];
        
        // Find all shape elements that can contain text
        const shapeMatches = slideContent.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
        
        for (const shape of shapeMatches) {
          // Check if this is a title shape - if so, we already processed it
          const isTitleShape = shape.includes('type="title"') || shape.includes('idx="0"');
          
          // Skip the title shape as we already extracted it
          if (isTitleShape && title && shape.includes(title)) {
            continue;
          }
          
          // Extract text from this shape
          const textMatches = shape.match(/<a:t>[\s\S]*?<\/a:t>/g) || [];
          if (textMatches.length > 0) {
            const shapeText = textMatches.map(match => 
              match.replace(/<a:t>([\s\S]*?)<\/a:t>/, '$1')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&amp;/g, '&')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'")
            ).join('\n');
            
            if (shapeText.trim()) {
              structuredContent.push({
                type: 'text',
                content: shapeText
              });
            }
          }
        }
        
        // Extract all text content for backward compatibility
        const allTextMatches = slideContent.match(/<a:t>[\s\S]*?<\/a:t>/g) || [];
        const allTexts = allTextMatches.map(match => 
          match.replace(/<a:t>([\s\S]*?)<\/a:t>/, '$1')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
        );
        
        // Extract table data with improved structure preservation
        const tables = [];
        const tableMatches = slideContent.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) || [];
        
        tableMatches.forEach(tableMatch => {
          // Get table grid information - this defines columns
          const gridColsMatch = tableMatch.match(/<a:gridCol[^>]*\/>/g) || [];
          const colsCount = gridColsMatch.length;
          
          // Analyse grid column widths for better proportions
          const colWidths = gridColsMatch.map(col => {
            const widthMatch = col.match(/w="([^"]+)"/);
            return widthMatch ? parseInt(widthMatch[1]) : 0;
          });
          
          // Extract rows
          const rowsMatch = tableMatch.match(/<a:tr[^>]*>[\s\S]*?<\/a:tr>/g) || [];
          const tableRows = [];
          
          let hasHeader = false;
          let headerStyle = null;
          
          // Determine the maximum number of columns across all rows
          let maxColsInTable = 0;
          rowsMatch.forEach(rowMatch => {
            const cellsCount = (rowMatch.match(/<a:tc/g) || []).length;
            const gridSpans = rowMatch.match(/gridSpan="([^"]+)"/g) || [];
            let effectiveColCount = cellsCount;
            
            // Add grid spans to effective column count
            for (const span of gridSpans) {
              const spanValue = parseInt(span.match(/gridSpan="([^"]+)"/)[1]) - 1;
              effectiveColCount += spanValue;
            }
            
            maxColsInTable = Math.max(maxColsInTable, effectiveColCount);
          });
          
          // If we don't have column count from gridCol elements, use max cols found in rows
          const totalColsCount = colsCount || maxColsInTable;
          
          rowsMatch.forEach((rowMatch, rowIndex) => {
            // Check for header styling
            if (rowIndex === 0) {
              const boldFormat = rowMatch.includes('<a:b/>') || rowMatch.includes('i="bold"');
              const bgFillColor = rowMatch.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
              
              if (boldFormat || bgFillColor) {
                hasHeader = true;
                headerStyle = {
                  bold: boldFormat,
                  bgColor: bgFillColor ? bgFillColor[1] : null
                };
              }
            }
            
            // Extract cells in this row
            const cellsMatch = rowMatch.match(/<a:tc>[\s\S]*?<\/a:tc>/g) || [];
            const rowCells = [];
            
            let currentColIndex = 0;
            
            for (let i = 0; i < cellsMatch.length; i++) {
              const cellMatch = cellsMatch[i];
              
              // Check for row/column spans
              const rowSpanMatch = cellMatch.match(/rowSpan="([^"]+)"/);
              const colSpanMatch = cellMatch.match(/gridSpan="([^"]+)"/);
              
              const rowSpan = rowSpanMatch ? parseInt(rowSpanMatch[1]) : 1;
              const colSpan = colSpanMatch ? parseInt(colSpanMatch[1]) : 1;
              
              // Get cell text content
              const cellTextMatch = cellMatch.match(/<a:t>[\s\S]*?<\/a:t>/g) || [];
              const cellText = cellTextMatch.map(t => 
                t.replace(/<a:t>([\s\S]*?)<\/a:t>/, '$1')
                  .replace(/&lt;/g, '<')
                  .replace(/&gt;/g, '>')
                  .replace(/&amp;/g, '&')
                  .replace(/&quot;/g, '"')
                  .replace(/&apos;/g, "'")
              ).join(' ');
              
              // Get cell styling
              const isBold = cellMatch.includes('<a:b/>') || cellMatch.includes('i="bold"');
              const isItalic = cellMatch.includes('<a:i/>') || cellMatch.includes('i="italic"');
              const isUnderline = cellMatch.includes('<a:u/>') || cellMatch.includes('i="underline"');
              const bgFillColorMatch = cellMatch.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
              
              // Create cell object with all attributes
              rowCells.push({
                text: cellText,
                rowSpan,
                colSpan,
                colIndex: currentColIndex,
                style: {
                  bold: isBold,
                  italic: isItalic,
                  underline: isUnderline,
                  bgColor: bgFillColorMatch ? bgFillColorMatch[1] : null
                }
              });
              
              // Update column index for next cell
              currentColIndex += colSpan;
            }
            
            // Fill in missing cells to match column count
            while (currentColIndex < totalColsCount) {
              rowCells.push({
                text: '',
                rowSpan: 1,
                colSpan: 1,
                colIndex: currentColIndex,
                style: {}
              });
              currentColIndex++;
            }
            
            tableRows.push(rowCells);
          });
          
          tables.push({
            rows: tableRows,
            hasHeader,
            headerStyle,
            colsCount: totalColsCount,
            colWidths: colWidths.length > 0 ? colWidths : null
          });
          
          // Add table to structured content
          structuredContent.push({
            type: 'table',
            content: tableRows.map(row => row.map(cell => cell.text)),
            tableObject: {
              rows: tableRows,
              hasHeader,
              headerStyle,
              colsCount: totalColsCount,
              colWidths: colWidths.length > 0 ? colWidths : null
            }
          });
        });
        
        // Extract images (just references at this point)
        const images = [];
        if (slideRelsXml) {
          const imageRelMatches = slideRelsXml.match(/Relationship[^>]*Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/image"[^>]*Target="([^"]+)"/g) || [];
          
          imageRelMatches.forEach(relMatch => {
            const targetMatch = relMatch.match(/Target="([^"]+)"/);
            if (targetMatch) {
              const imagePath = targetMatch[1].startsWith('../') 
                ? `ppt/${targetMatch[1].substring(3)}` 
                : `ppt/slides/${targetMatch[1]}`;
              
              images.push(imagePath);
              
              // Add image to structured content
              structuredContent.push({
                type: 'image',
                path: imagePath
              });
            }
          });
        }
        
        // Construct a structured representation of slide content
        const slideData = {
          id: `slide-${slideId}`,
          title: title || `Slide ${slideId}`,
          structuredContent,
          texts: allTexts,
          content: allTexts.join('\n'),
          tables,
          images,
          rawXml: slideContent // Store raw XML for advanced users
        };
        
        parsedSlides.push(slideData);
      }
      
      // Extract images for slide previews
      for (let slide of parsedSlides) {
        // Process all image paths to get urls
        const imageElements = slide.structuredContent.filter(item => item.type === 'image');
        
        for (let i = 0; i < slide.images.length; i++) {
          const imagePath = slide.images[i];
          if (zipData.files[imagePath]) {
            try {
              const imageData = await zipData.files[imagePath].async('blob');
              const imageUrl = URL.createObjectURL(imageData);
              slide.images[i] = { path: imagePath, url: imageUrl };
              
              // Update structured content image references too
              for (let element of imageElements) {
                if (element.path === imagePath) {
                  element.url = imageUrl;
                }
              }
            } catch (err) {
              console.error(`Error extracting image ${imagePath}:`, err);
              slide.images[i] = { path: imagePath, error: err.message };
            }
          }
        }
      }
      
      return parsedSlides.length > 0 ? parsedSlides : [{ 
        id: 'slide-1',
        title: 'New Presentation', 
        content: 'Click to add content',
        texts: ['Click to add content'],
        structuredContent: [{ type: 'text', content: 'Click to add content' }],
        tables: [],
        images: []
      }];
    } catch (error) {
      console.error('Error parsing PPTX:', error);
      // Return a default slide if parsing fails
      return [{ 
        id: 'slide-1',
        title: 'New Presentation', 
        content: 'Click to add content',
        texts: ['Click to add content'],
        structuredContent: [{ type: 'text', content: 'Click to add content' }],
        tables: [],
        images: []
      }];
    }
  };

  // Handle title change
  const handleTitleChange = (e) => {
    const updatedSlides = [...slides];
    updatedSlides[selectedSlide].title = e.target.value;
    setSlides(updatedSlides);
  };

  // Handle content change
  const handleContentChange = (e) => {
    const updatedSlides = [...slides];
    updatedSlides[selectedSlide].content = e.target.value;
    // Also update the texts array
    updatedSlides[selectedSlide].texts = [e.target.value];
    setSlides(updatedSlides);
  };

  // Generate a preview HTML for slide thumbnails
  const generateSlidePreview = (slide) => {
    let previewContent = `<div class="slide-preview-content">`;
    
    // Add title - always at the top with emphasis
    previewContent += `<div class="preview-title">${slide.title}</div>`;
    
    // Add structured content in order
    if (slide.structuredContent && slide.structuredContent.length > 0) {
      slide.structuredContent.forEach(item => {
        if (item.type === 'text') {
          previewContent += `<div class="preview-text">${item.content}</div>`;
        } else if (item.type === 'table') {
          previewContent += `<table class="preview-table">`;
          
          // Add table with proper structure
          const tableObj = item.tableObject || { rows: item.content.map(row => row.map(cell => ({ text: cell }))) };
          
          tableObj.rows.forEach((row, rowIndex) => {
            previewContent += `<tr ${rowIndex === 0 && tableObj.hasHeader ? 'class="header-row"' : ''}>`;
            row.forEach(cell => {
              const style = cell.style ? 
                `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
              
              previewContent += `<td ${style} ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}>${cell.text || ''}</td>`;
            });
            previewContent += `</tr>`;
          });
          
          previewContent += `</table>`;
        } else if (item.type === 'image' && item.url) {
          previewContent += `<img src="${item.url}" class="preview-image" alt="Slide image" />`;
        }
      });
    } else {
      // Fallback to simple text display
      if (slide.texts && slide.texts.length > 1) {
        previewContent += `<div class="preview-text">${slide.texts.slice(1).join('<br/>')}</div>`;
      }
      
      // Add tables
      if (slide.tables && slide.tables.length > 0) {
        slide.tables.forEach(table => {
          previewContent += `<table class="preview-table">`;
          if (table.rows) {
            table.rows.forEach((row, i) => {
              previewContent += `<tr ${i === 0 && table.hasHeader ? 'class="header-row"' : ''}>`;
              row.forEach(cell => {
                const cellText = typeof cell === 'object' ? cell.text : cell;
                const style = typeof cell === 'object' && cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                previewContent += `<td ${style}>${cellText || ''}</td>`;
              });
              previewContent += `</tr>`;
            });
          } else if (Array.isArray(table)) {
            table.forEach(row => {
              previewContent += `<tr>`;
              row.forEach(cell => {
                previewContent += `<td>${cell}</td>`;
              });
              previewContent += `</tr>`;
            });
          }
          previewContent += `</table>`;
        });
      }
      
      // Add images
      if (slide.images && slide.images.length > 0) {
        slide.images.forEach(image => {
          if (image.url) {
            previewContent += `<img src="${image.url}" class="preview-image" alt="Slide image" />`;
          }
        });
      }
    }
    
    previewContent += `</div>`;
    return previewContent;
  };

  // Generate a preview HTML for slide thumbnails with editable elements
  const generateEditableSlidePreview = (slide, slideIndex) => {
    let previewContent = `<div class="slide-preview-content">`;
    
    // Add title - always at the top with emphasis, now editable
    previewContent += `<div 
      class="preview-title editable" 
      contenteditable="true"
      data-slide-index="${slideIndex}" 
      data-element-type="title"
    >${slide.title}</div>`;
    
    // Add structured content in order, with editable text
    if (slide.structuredContent && slide.structuredContent.length > 0) {
      slide.structuredContent.forEach((item, itemIndex) => {
        if (item.type === 'text') {
          previewContent += `<div 
            class="preview-text editable" 
            contenteditable="true"
            data-slide-index="${slideIndex}" 
            data-element-type="text"
            data-element-index="${itemIndex}"
          >${item.content}</div>`;
        } else if (item.type === 'table') {
          previewContent += `<div class="table-container">`;
          previewContent += `<table class="preview-table">`;
          
          // Add table with proper structure and editable cells
          const tableObj = item.tableObject || { rows: item.content.map(row => row.map(cell => ({ text: cell }))) };
          
          // Check if we should render based on column metrics
          let styleAttr = '';
          if (tableObj.colWidths && tableObj.colWidths.length > 0) {
            // Calculate column proportions based on width values
            const totalWidth = tableObj.colWidths.reduce((a, b) => a + b, 0);
            if (totalWidth > 0) {
              previewContent += `<colgroup>`;
              tableObj.colWidths.forEach(width => {
                const percentage = Math.round((width / totalWidth) * 100);
                previewContent += `<col style="width: ${percentage}%">`;
              });
              previewContent += `</colgroup>`;
            }
          }
          
          // Create header row if needed with correct column count
          if (tableObj.hasHeader && tableObj.rows && tableObj.rows.length > 0) {
            const headerRow = tableObj.rows[0];
            previewContent += `<thead><tr class="header-row">`;
            
            // Go through header cells and apply colspan as needed
            headerRow.forEach(cell => {
              const style = cell.style ? 
                `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
              
              previewContent += `<th 
                ${style} 
                ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} 
                ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}
                contenteditable="true"
                class="editable"
                data-slide-index="${slideIndex}"
                data-element-type="table"
                data-element-index="${itemIndex}"
                data-row-index="0"
                data-cell-index="${cell.colIndex || 0}"
              >${cell.text || ''}</th>`;
            });
            
            previewContent += `</tr></thead><tbody>`;
            
            // Render body rows starting from row 1
            tableObj.rows.slice(1).forEach((row, rowIndexOffset) => {
              const rowIndex = rowIndexOffset + 1; // actual index in the rows array
              previewContent += `<tr>`;
              
              row.forEach(cell => {
                const style = cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                previewContent += `<td 
                  ${style} 
                  ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} 
                  ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}
                  contenteditable="true"
                  class="editable"
                  data-slide-index="${slideIndex}"
                  data-element-type="table"
                  data-element-index="${itemIndex}"
                  data-row-index="${rowIndex}"
                  data-cell-index="${cell.colIndex || 0}"
                >${cell.text || ''}</td>`;
              });
              
              previewContent += `</tr>`;
            });
            
            previewContent += `</tbody>`;
          } else {
            // No specific header, render all rows normally
            previewContent += `<tbody>`;
            
            tableObj.rows.forEach((row, rowIndex) => {
              previewContent += `<tr>`;
              
              row.forEach(cell => {
                const style = cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                previewContent += `<td 
                  ${style} 
                  ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} 
                  ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}
                  contenteditable="true"
                  class="editable"
                  data-slide-index="${slideIndex}"
                  data-element-type="table"
                  data-element-index="${itemIndex}"
                  data-row-index="${rowIndex}"
                  data-cell-index="${cell.colIndex || 0}"
                >${cell.text || ''}</td>`;
              });
              
              previewContent += `</tr>`;
            });
            
            previewContent += `</tbody>`;
          }
          
          previewContent += `</table>`;
          previewContent += `</div>`;
        } else if (item.type === 'image' && item.url) {
          previewContent += `<img src="${item.url}" class="preview-image" alt="Slide image" />`;
        }
      });
    } else {
      // Fallback to simple text display, now editable
      if (slide.texts && slide.texts.length > 1) {
        slide.texts.slice(1).forEach((text, textIndex) => {
          previewContent += `<div 
            class="preview-text editable" 
            contenteditable="true"
            data-slide-index="${slideIndex}" 
            data-element-type="legacy-text"
            data-element-index="${textIndex + 1}"
          >${text}</div>`;
        });
      }
      
      // Add tables with editable cells
      if (slide.tables && slide.tables.length > 0) {
        slide.tables.forEach((table, tableIndex) => {
          previewContent += `<table class="preview-table">`;
          if (table.rows) {
            table.rows.forEach((row, rowIndex) => {
              previewContent += `<tr ${rowIndex === 0 && table.hasHeader ? 'class="header-row"' : ''}>`;
              row.forEach((cell, cellIndex) => {
                const cellText = typeof cell === 'object' ? cell.text : cell;
                const style = typeof cell === 'object' && cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                previewContent += `<td 
                  ${style}
                  contenteditable="true"
                  class="editable"
                  data-slide-index="${slideIndex}"
                  data-element-type="legacy-table"
                  data-table-index="${tableIndex}"
                  data-row-index="${rowIndex}"
                  data-cell-index="${cellIndex}"
                >${cellText || ''}</td>`;
              });
              previewContent += `</tr>`;
            });
          } else if (Array.isArray(table)) {
            table.forEach((row, rowIndex) => {
              previewContent += `<tr>`;
              row.forEach((cell, cellIndex) => {
                previewContent += `<td
                  contenteditable="true"
                  class="editable"
                  data-slide-index="${slideIndex}"
                  data-element-type="legacy-table-simple"
                  data-table-index="${tableIndex}"
                  data-row-index="${rowIndex}"
                  data-cell-index="${cellIndex}"
                >${cell}</td>`;
              });
              previewContent += `</tr>`;
            });
          }
          previewContent += `</table>`;
        });
      }
      
      // Add images
      if (slide.images && slide.images.length > 0) {
        slide.images.forEach(image => {
          if (image.url) {
            previewContent += `<img src="${image.url}" class="preview-image" alt="Slide image" />`;
          }
        });
      }
    }
    
    previewContent += `</div>`;
    return previewContent;
  };

  // Handle content edit in the structured preview
  const handleStructuredContentEdit = (event) => {
    const element = event.target;
    const slideIndex = parseInt(element.getAttribute('data-slide-index'));
    const elementType = element.getAttribute('data-element-type');
    const updatedContent = element.innerText;
    
    if (isNaN(slideIndex) || !elementType) return;
    
    const updatedSlides = [...slides];
    const slide = updatedSlides[slideIndex];
    
    if (elementType === 'title') {
      // Update title
      slide.title = updatedContent;
      // Also update the textarea if this is the selected slide
      if (slideIndex === selectedSlide) {
        document.querySelector('.slide-title-input').value = updatedContent;
      }
    } 
    else if (elementType === 'text') {
      // Update text in structured content
      const elementIndex = parseInt(element.getAttribute('data-element-index'));
      if (!isNaN(elementIndex) && slide.structuredContent && slide.structuredContent[elementIndex]) {
        slide.structuredContent[elementIndex].content = updatedContent;
        
        // Update the content textarea if this is the selected slide
        if (slideIndex === selectedSlide) {
          // Recreate content text from all structured content
          const textContent = slide.structuredContent
            .filter(item => item.type === 'text')
            .map(item => item.content)
            .join('\n\n');
          
          slide.content = textContent;
          // Update the textarea if it exists
          const textarea = document.querySelector('.slide-content-input');
          if (textarea) {
            textarea.value = textContent;
          }
        }
      }
    }
    else if (elementType === 'legacy-text') {
      // Update text in the texts array
      const elementIndex = parseInt(element.getAttribute('data-element-index'));
      if (!isNaN(elementIndex) && slide.texts && slide.texts[elementIndex]) {
        slide.texts[elementIndex] = updatedContent;
        
        // Update the content textarea if this is the selected slide
        if (slideIndex === selectedSlide) {
          slide.content = slide.texts.join('\n\n');
          const textarea = document.querySelector('.slide-content-input');
          if (textarea) {
            textarea.value = slide.content;
          }
        }
      }
    }
    else if (elementType.includes('table')) {
      // Update cell in table
      const elementIndex = parseInt(element.getAttribute('data-element-index'));
      const tableIndex = parseInt(element.getAttribute('data-table-index'));
      const rowIndex = parseInt(element.getAttribute('data-row-index'));
      const cellIndex = parseInt(element.getAttribute('data-cell-index'));
      
      let targetTable;
      
      if (elementType === 'table' && !isNaN(elementIndex) && slide.structuredContent) {
        // Table in structured content
        targetTable = slide.structuredContent[elementIndex]?.tableObject;
        
        if (targetTable && targetTable.rows && targetTable.rows[rowIndex] && targetTable.rows[rowIndex][cellIndex]) {
          targetTable.rows[rowIndex][cellIndex].text = updatedContent;
          
          // Also update the content array for backward compatibility
          if (slide.structuredContent[elementIndex].content) {
            slide.structuredContent[elementIndex].content[rowIndex][cellIndex] = updatedContent;
          }
        }
      } 
      else if ((elementType === 'legacy-table' || elementType === 'legacy-table-simple') && !isNaN(tableIndex)) {
        // Legacy table format
        targetTable = slide.tables[tableIndex];
        
        if (targetTable) {
          if (elementType === 'legacy-table' && targetTable.rows && targetTable.rows[rowIndex] && targetTable.rows[rowIndex][cellIndex]) {
            // Complex table with row objects
            if (typeof targetTable.rows[rowIndex][cellIndex] === 'object') {
              targetTable.rows[rowIndex][cellIndex].text = updatedContent;
            } else {
              targetTable.rows[rowIndex][cellIndex] = updatedContent;
            }
          } 
          else if (elementType === 'legacy-table-simple' && Array.isArray(targetTable) && targetTable[rowIndex] && targetTable[rowIndex][cellIndex]) {
            // Simple array table
            targetTable[rowIndex][cellIndex] = updatedContent;
          }
        }
      }
    }
    
    // Update the slides state
    setSlides(updatedSlides);
  };

  // Generate detailed table preview for the editor
  const generateTablePreview = (table) => {
    let tableHtml = `<table class="detailed-table">`;
    
    // Add column group if we have column widths
    if (table.colWidths && table.colWidths.length > 0) {
      const totalWidth = table.colWidths.reduce((a, b) => a + b, 0);
      if (totalWidth > 0) {
        tableHtml += `<colgroup>`;
        table.colWidths.forEach(width => {
          const percentage = Math.round((width / totalWidth) * 100);
          tableHtml += `<col style="width: ${percentage}%">`;
        });
        tableHtml += `</colgroup>`;
      }
    }
    
    // Determine if we have a header row
    if (table.hasHeader && table.rows && table.rows.length > 0) {
      const headerRow = table.rows[0];
      
      tableHtml += `<thead><tr class="header-row">`;
      
      headerRow.forEach(cell => {
        const style = cell.style ? 
          `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
        
        tableHtml += `<th ${style} ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}>${cell.text || ''}</th>`;
      });
      
      tableHtml += `</tr></thead><tbody>`;
      
      // Render body rows starting from row 1
      table.rows.slice(1).forEach(row => {
        tableHtml += `<tr>`;
        
        row.forEach(cell => {
          const style = cell.style ? 
            `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
          
          tableHtml += `<td ${style} ${cell.rowSpan > 1 ? `rowspan="${cell.rowSpan}"` : ''} ${cell.colSpan > 1 ? `colspan="${cell.colSpan}"` : ''}>${cell.text || ''}</td>`;
        });
        
        tableHtml += `</tr>`;
      });
      
      tableHtml += `</tbody>`;
    } else if (table.rows) {
      // No specific header, render all rows normally
      tableHtml += `<tbody>`;
      
      table.rows.forEach(row => {
        tableHtml += `<tr>`;
        
        row.forEach(cell => {
          const cellContent = typeof cell === 'object' ? cell.text : cell;
          const style = typeof cell === 'object' && cell.style ? 
            `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
          const rowSpan = typeof cell === 'object' && cell.rowSpan ? `rowspan="${cell.rowSpan}"` : '';
          const colSpan = typeof cell === 'object' && cell.colSpan ? `colspan="${cell.colSpan}"` : '';
          
          tableHtml += `<td ${style} ${rowSpan} ${colSpan}>${cellContent || ''}</td>`;
        });
        
        tableHtml += `</tr>`;
      });
      
      tableHtml += `</tbody>`;
    } else if (Array.isArray(table)) {
      // Simple array table format
      tableHtml += `<tbody>`;
      
      table.forEach(row => {
        tableHtml += `<tr>`;
        
        row.forEach(cell => {
          const cellContent = typeof cell === 'object' ? cell.text : cell;
          tableHtml += `<td>${cellContent || ''}</td>`;
        });
        
        tableHtml += `</tr>`;
      });
      
      tableHtml += `</tbody>`;
    }
    
    tableHtml += `</table>`;
    return tableHtml;
  };

  // Add a new slide
  const addSlide = () => {
    const newSlides = [...slides];
    newSlides.push({
      id: `slide-${newSlides.length + 1}`,
      title: 'New Slide',
      content: 'Click to add content',
      texts: ['Click to add content'],
      tables: [],
      images: []
    });
    setSlides(newSlides);
    setSelectedSlide(newSlides.length - 1);
  };

  // Delete the current slide
  const deleteSlide = () => {
    if (slides.length <= 1) {
      alert('Cannot delete the only slide');
      return;
    }

    const newSlides = slides.filter((_, index) => index !== selectedSlide);
    setSlides(newSlides);
    setSelectedSlide(Math.min(selectedSlide, newSlides.length - 1));
  };

  // Export PPTX
  const exportPptx = () => {
    try {
      const pptx = new pptxgen();
      
      // Create a slide for each slide in our state
      slides.forEach(slide => {
        const pptxSlide = pptx.addSlide();
        
        // Add title
        pptxSlide.addText(slide.title, { 
          x: 1, 
          y: 0.5, 
          fontSize: 24,
          bold: true
        });
        
        // Add content text
        if (slide.texts && slide.texts.length > 0) {
          pptxSlide.addText(slide.content, {
            x: 1,
            y: 1.5,
            fontSize: 14
          });
        }
        
        // Add tables if present
        if (slide.tables && slide.tables.length > 0) {
          slide.tables.forEach((table, index) => {
            const tableData = [];
            table.forEach(row => {
              tableData.push(row);
            });
            
            pptxSlide.addTable(tableData, {
              x: 1,
              y: 3 + (index * 2),
              w: 8,
              border: { pt: 1, color: '666666' }
            });
          });
        }
      });
      
      // Save the PPTX file
      const exportFileName = fileName ? fileName : 'presentation.pptx';
      pptx.writeFile({ fileName: exportFileName });
    } catch (error) {
      console.error('Error exporting PPTX:', error);
      alert('Failed to export the presentation');
    }
  };

  // Create new presentation from scratch
  const createNewPresentation = () => {
    setSlides([{
      id: 'slide-1',
      title: 'New Presentation',
      content: 'Click to add content',
      texts: ['Click to add content'],
      tables: [],
      images: []
    }]);
    setSelectedSlide(0);
    setFileName('new-presentation.pptx');
  };

  // When slide is selected, update the preview
  useEffect(() => {
    if (slides.length > 0 && previewRefs.current) {
      slides.forEach((slide, index) => {
        if (previewRefs.current[index]) {
          previewRefs.current[index].innerHTML = generateSlidePreview(slide);
        }
      });
    }
  }, [slides, selectedSlide]);
  
  // Add event listeners for editable content after the structured preview is rendered
  useEffect(() => {
    const structuredPreview = document.querySelector('.slide-structured-preview');
    if (structuredPreview) {
      // Remove existing listeners first to prevent duplicates
      const editableElements = structuredPreview.querySelectorAll('.editable');
      editableElements.forEach(element => {
        element.removeEventListener('blur', handleStructuredContentEdit);
      });
      
      // Add listeners to new elements
      editableElements.forEach(element => {
        element.addEventListener('blur', handleStructuredContentEdit);
      });
      
      // Cleanup when component unmounts
      return () => {
        editableElements.forEach(element => {
          element.removeEventListener('blur', handleStructuredContentEdit);
        });
      };
    }
  }, [selectedSlide, slides]);

  return (
    <div className="pptx-editor">
      {slides.length === 0 ? (
        <div className="dropzone-container" {...getRootProps()}>
          <input {...getInputProps()} />
          <div className="dropzone">
            {isDragActive ? (
              <p>Déposez le fichier PPTX ici...</p>
            ) : (
              <div>
                <p>Glissez-déposez un fichier PPTX ici, ou cliquez pour sélectionner un fichier</p>
                <button className="new-btn" onClick={(e) => { e.stopPropagation(); createNewPresentation(); }}>
                  Créer une nouvelle présentation
                </button>
              </div>
            )}
          </div>
        </div>
      ) : (
        <div className="editor-container">
          <div className="editor-header">
            <h2>{fileName || 'Présentation sans titre'}</h2>
            <div className="editor-actions">
              <button onClick={addSlide}>Ajouter une diapositive</button>
              <button onClick={deleteSlide}>Supprimer la diapositive</button>
              <button onClick={exportPptx}>Exporter en PPTX</button>
              <button onClick={() => { setSlides([]); setFileName(''); }}>
                Charger un autre fichier
              </button>
            </div>
          </div>
          
          <div className="editor-content">
            <div className="slides-thumbnail">
              {slides.map((slide, index) => (
                <div 
                  key={slide.id}
                  className={`slide-thumb ${selectedSlide === index ? 'selected' : ''}`}
                  onClick={() => setSelectedSlide(index)}
                >
                  <div className="slide-number">{index + 1}</div>
                  <div 
                    className="slide-preview-container"
                    ref={el => previewRefs.current[index] = el}
                  ></div>
                </div>
              ))}
            </div>
            
            <div className="slide-editor">
              {isLoading ? (
                <div className="loading">Chargement de la présentation...</div>
              ) : (
                <>
                  <input
                    type="text"
                    value={slides[selectedSlide].title}
                    onChange={handleTitleChange}
                    className="slide-title-input"
                  />
                  
                  <div className="slide-structure-container">
                    <h3>Aperçu structuré <span className="edit-hint">(cliquez sur les éléments pour les modifier)</span></h3>
                    <div className="structured-preview">
                      <div 
                        dangerouslySetInnerHTML={{ __html: generateEditableSlidePreview(slides[selectedSlide], selectedSlide) }} 
                        className="slide-structured-preview"
                      />
                    </div>
                  </div>
                  
                  <textarea
                    value={slides[selectedSlide].content}
                    onChange={handleContentChange}
                    className="slide-content-input"
                    placeholder="Éditez le contenu ici ou directement dans l'aperçu structuré ci-dessus"
                  />
                  
                  {slides[selectedSlide].tables && slides[selectedSlide].tables.length > 0 && (
                    <div className="tables-container">
                      <h3>Tableaux dans cette diapositive</h3>
                      {slides[selectedSlide].tables.map((table, tableIndex) => (
                        <div key={`table-${tableIndex}`} className="table-preview">
                          <div dangerouslySetInnerHTML={{ __html: generateTablePreview(table) }} />
                        </div>
                      ))}
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PptxViewer;