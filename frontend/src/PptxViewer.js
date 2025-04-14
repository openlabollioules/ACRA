// PptxViewer.js utilise our api to get the pptx file strcutre then display it in a react component
import React, { useState, useCallback, useEffect, useRef } from 'react';
import { useDropzone } from 'react-dropzone';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import axios from 'axios';
import './PptxViewer.css';

// Déterminer l'URL de base de l'API en fonction de l'environnement
const API_BASE_URL = process.env.NODE_ENV === 'production' 
  ? '/api' 
  : window.location.hostname === 'localhost' 
    ? 'http://localhost:5050'
    : 'http://fastapi:5050';

// Composant de popup pour le formatage du texte
const TextFormatPopup = ({ position, onFormat, onClose }) => {
  const [dragPosition, setDragPosition] = React.useState({ x: position.left, y: position.top });
  const [isDragging, setIsDragging] = React.useState(false);
  const [dragOffset, setDragOffset] = React.useState({ x: 0, y: 0 });
  const popupRef = React.useRef(null);

  // Gestionnaire de début de déplacement
  const handleMouseDown = (e) => {
    // Ne déclencher que si le clic est sur la barre du popup, pas sur les boutons
    if (e.target.tagName !== 'BUTTON') {
      setIsDragging(true);
      const rect = popupRef.current.getBoundingClientRect();
      setDragOffset({
        x: e.clientX - rect.left,
        y: e.clientY - rect.top
      });
      e.preventDefault();
    }
  };

  // Gestionnaire de déplacement
  const handleMouseMove = React.useCallback((e) => {
    if (isDragging) {
      const newX = e.clientX - dragOffset.x;
      const newY = e.clientY - dragOffset.y;
      
      setDragPosition({ x: newX, y: newY });
      
      // Enregistrer la position dans localStorage
      localStorage.setItem('formatPopupPosition', JSON.stringify({ x: newX, y: newY }));
    }
  }, [isDragging, dragOffset]);

  // Gestionnaire de fin de déplacement
  const handleMouseUp = React.useCallback(() => {
    if (isDragging) {
      setIsDragging(false);
    }
  }, [isDragging]);

  // Ajouter/retirer les listeners de déplacement
  React.useEffect(() => {
    if (isDragging) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
    } else {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    }
    
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDragging, handleMouseMove, handleMouseUp]);

  const popupStyle = {
    position: 'fixed', // 'fixed' au lieu de 'absolute' pour éviter le décalage lors du scroll
    top: `${isDragging ? dragPosition.y : position.top}px`,
    left: `${isDragging ? dragPosition.x : position.left}px`,
    transform: isDragging ? 'none' : 'translate(-50%, -100%)',
    backgroundColor: 'white',
    border: '1px solid #ccc',
    borderRadius: '4px',
    padding: '5px',
    boxShadow: '0 2px 5px rgba(0,0,0,0.2)',
    zIndex: 9999,
    display: 'flex',
    alignItems: 'center',
    cursor: isDragging ? 'grabbing' : 'grab'
  };

  const buttonStyle = {
    margin: '0 5px',
    padding: '3px 8px',
    cursor: 'pointer',
    border: 'none',
    borderRadius: '3px'
  };

  return (
    <div 
      ref={popupRef}
      style={popupStyle} 
      className={`format-popup ${isDragging ? 'dragging' : ''}`}
      onMouseDown={handleMouseDown}
    >
      <button 
        style={{ ...buttonStyle, color: 'black', backgroundColor: '#f0f0f0' }} 
        onClick={() => onFormat('black')}
      >
        A
      </button>
      <button 
        style={{ ...buttonStyle, color: 'green', backgroundColor: '#f0f0f0' }} 
        onClick={() => onFormat('green')}
      >
        A
      </button>
      <button 
        style={{ ...buttonStyle, color: 'red', backgroundColor: '#f0f0f0' }} 
        onClick={() => onFormat('red')}
      >
        A
      </button>
      <button 
        style={{ ...buttonStyle, fontStyle: 'italic', backgroundColor: '#f0f0f0' }} 
        onClick={() => onFormat('italic')}
      >
        I
      </button>
      <button 
        style={{ ...buttonStyle, backgroundColor: '#f0f0f0' }} 
        onClick={onClose}
      >
        ✕
      </button>
    </div>
  );
};

const PptxViewer = ({ chatId }) => {
  const [presentations, setPresentations] = useState([]);
  const [currentPresIndex, setCurrentPresIndex] = useState(0);
  const [selectedSlide, setSelectedSlide] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [folderName, setFolderName] = useState('');
  const [formatPopup, setFormatPopup] = useState({ visible: false, position: { top: 0, left: 0 } });
  const [currentSelection, setCurrentSelection] = useState(null);
  const previewRefs = useRef([]);
  const previousChatId = useRef(null);

  // Fermer le popup de formatage avec useCallback - définir en premier pour éviter les erreurs de référence
  const closeFormatPopup = useCallback(() => {
    setFormatPopup({ visible: false, position: { top: 0, left: 0 } });
    setCurrentSelection(null);
    window.getSelection().removeAllRanges();
  }, []);

  // Gestionnaire de clic en dehors du popup pour le fermer avec useCallback
  const handleClickOutside = useCallback((event) => {
    if (formatPopup.visible && !event.target.closest('.format-popup') && 
        !event.target.closest('.editable')) {
      closeFormatPopup();
    }
  }, [formatPopup.visible, closeFormatPopup]);

  // Extraire les alertes du contenu HTML formaté
  const extractAlertsFromHtml = useCallback((htmlContent) => {
    const alerts = {
      advancements: [],
      small_alerts: [],
      critical_alerts: []
    };
    
    // Extraire les avancées (vert)
    const advancementMatches = htmlContent.match(/<span style="color: green;">(.*?)<\/span>/g);
    if (advancementMatches) {
      advancementMatches.forEach(match => {
        const content = match.replace(/<span style="color: green;">/, '').replace(/<\/span>/, '');
        alerts.advancements.push(content);
      });
    }
    
    // Extraire les alertes (orange)
    const smallAlertMatches = htmlContent.match(/<span style="color: orange;">(.*?)<\/span>/g);
    if (smallAlertMatches) {
      smallAlertMatches.forEach(match => {
        const content = match.replace(/<span style="color: orange;">/, '').replace(/<\/span>/, '');
        alerts.small_alerts.push(content);
      });
    }
    
    // Extraire les alertes critiques (rouge)
    const criticalAlertMatches = htmlContent.match(/<span style="color: red; font-weight: bold;">(.*?)<\/span>/g);
    if (criticalAlertMatches) {
      criticalAlertMatches.forEach(match => {
        const content = match.replace(/<span style="color: red; font-weight: bold;">/, '').replace(/<\/span>/, '');
        alerts.critical_alerts.push(content);
      });
    }
    
    return alerts;
  }, []);

  // Get presentations from API
  const fetchPresentations = async (folder = chatId) => {
    if (!folder) {
      console.log('Aucun dossier/chatId spécifié, aucune présentation ne sera chargée');
      setPresentations([]);
      setError('Sélectionnez une conversation pour afficher les présentations');
      return [];
    }
    
    setIsLoading(true);
    setError(null);
    try {
      console.log(`Chargement des présentations pour le dossier/chatId: ${folder}`);
      const response = await axios.get(`${API_BASE_URL}/get_slide_structure/${folder}`);
      
      if (response.data && response.data.presentations && response.data.presentations.length > 0) {
        setPresentations(response.data.presentations);
        setCurrentPresIndex(0); // Reset to first presentation
        console.log(`${response.data.presentations.length} présentations chargées avec succès`);
        return response.data.presentations;
      } else {
        setError('Aucune présentation trouvée dans ce dossier');
        setPresentations([]);
        return [];
      }
    } catch (err) {
      console.error('Erreur lors de la récupération des présentations:', err);
      setError(`Erreur: ${err.response?.data?.detail || err.message}`);
      setPresentations([]);
      return [];
    } finally {
      setIsLoading(false);
    }
  };

  // Charger les présentations lorsque le chatId change
  useEffect(() => {
    if (chatId) {
      console.log(`ChatId disponible: ${chatId}`);
      setFolderName(chatId);
      
      if (chatId !== previousChatId.current) {
        console.log(`Nouveau chatId détecté: ${chatId}, différent de ${previousChatId.current}`);
        fetchPresentations(chatId);
        previousChatId.current = chatId;
      } else {
        console.log(`ChatId inchangé: ${chatId}, vérification si des présentations sont déjà chargées`);
        if (presentations.length === 0) {
          console.log(`Aucune présentation chargée, chargement forcé pour le chatId: ${chatId}`);
          fetchPresentations(chatId);
        }
      }
    } else {
      console.log('Aucun chatId disponible');
    }
  }, [chatId, presentations.length]);

  // Convert project_data to slides format for our editor
  const convertProjectDataToSlides = useCallback((projectData) => {
    const slides = [];
    
    // Create a single slide with a table containing all projects
    const mainSlide = {
      id: 'slide-1',
      title: projectData.metadata?.title || 'Présentation sans titre',
      content: '',
      structuredContent: [
        { type: 'text', content: projectData.metadata?.title || 'Présentation sans titre', isTitle: true }
      ],
      tables: [],
      images: []
    };
    
    // Create the project table if there are activities
    if (projectData.activities && Object.keys(projectData.activities).length > 0) {
      // Create table rows for each project
      const tableRows = [];
      
      // Header row
      tableRows.push([
        { text: 'Projet', style: { bold: true, bgColor: 'F0F0F0' }, rowSpan: 1, colSpan: 1 },
        { text: 'Information', style: { bold: true, bgColor: 'F0F0F0' }, rowSpan: 1, colSpan: 1 },
        { text: 'Événements à venir', style: { bold: true, bgColor: 'F0F0F0' }, rowSpan: 1, colSpan: 1 }
      ]);
      
      // Process each project as a row
      Object.entries(projectData.activities).forEach(([projectName, projectInfo], index) => {
        // Prepare information text with color-coded alerts
        let informationHtml = projectInfo.information || '';
        
        // Add color-coded alerts if they exist
        // Advancements (green)
        if (projectInfo.alerts?.advancements?.length > 0) {
          projectInfo.alerts.advancements.forEach(advancement => {
            informationHtml = informationHtml.replace(
              advancement,
              `<span style="color: green;">${advancement}</span>`
            );
          });
        }
        
        // Small alerts (orange)
        if (projectInfo.alerts?.small_alerts?.length > 0) {
          projectInfo.alerts.small_alerts.forEach(alert => {
            informationHtml = informationHtml.replace(
              alert,
              `<span style="color: orange;">${alert}</span>`
            );
          });
        }
        
        // Critical alerts (red)
        if (projectInfo.alerts?.critical_alerts?.length > 0) {
          projectInfo.alerts.critical_alerts.forEach(alert => {
            informationHtml = informationHtml.replace(
              alert,
              `<span style="color: red; font-weight: bold;">${alert}</span>`
            );
          });
        }
        
        // Create the row with project information
        const projectRow = [
          { text: projectName, style: { bold: true }, rowSpan: 1, colSpan: 1 },
          { text: informationHtml, html: true, rowSpan: 1, colSpan: 1 }
        ];
        
        // Add events column only to first row, spanning all project rows
        if (index === 0) {
          projectRow.push({
            text: projectData.upcoming_events || '',
            rowSpan: Object.keys(projectData.activities).length,
            colSpan: 1,
            style: { bgColor: 'F5F5F5' }
          });
        }
        
        tableRows.push(projectRow);
      });
      
      // Add the table to the slide
      mainSlide.tables.push({
        rows: tableRows,
        hasHeader: true,
        headerStyle: { bold: true, bgColor: 'F0F0F0' },
        colsCount: 3,
        colWidths: [150, 450, 250]
      });
      
      // Add the table to structured content
      mainSlide.structuredContent.push({
        type: 'table',
        content: tableRows.map(row => row.map(cell => cell.text)),
        tableObject: {
          rows: tableRows,
          hasHeader: true,
          headerStyle: { bold: true, bgColor: 'F0F0F0' },
          colsCount: 3,
          colWidths: [150, 450, 250]
        }
      });
    }
    
    // Add the main slide to the slides array
    slides.push(mainSlide);
    
    return slides;
  }, []);

  // Gestionnaire de sélection de texte avec useCallback
  const handleTextSelection = useCallback((event) => {
    const selection = window.getSelection();
    
    if (selection.toString().length > 0) {
      // Récupérer les informations sur la sélection
      const range = selection.getRangeAt(0);
      const rect = range.getBoundingClientRect();
      
      // Stocker la sélection actuelle
      setCurrentSelection({
        text: selection.toString(),
        range,
        element: event.target
      });
      
      // Déterminer la position du popup
      let position;
      
      // Vérifier s'il existe une position sauvegardée
      const savedPosition = localStorage.getItem('formatPopupPosition');
      
      if (savedPosition) {
        // Utiliser la position sauvegardée
        position = JSON.parse(savedPosition);
        
        // Vérifier si la position sauvegardée est dans la zone visible de l'écran
        const isVisible = 
          position.x > 0 && 
          position.x < window.innerWidth && 
          position.y > 0 && 
          position.y < window.innerHeight;
        
        // Si la position n'est pas visible, alors calculer la position par défaut
        if (!isVisible) {
          position = {
            left: rect.left + (rect.width / 2),
            top: rect.top + window.scrollY - 5
          };
        } else {
          // Convertir le format sauvegardé au format attendu par le composant
          position = {
            left: position.x,
            top: position.y
          };
        }
      } else {
        // Calculer la position par défaut (centrée au-dessus du texte sélectionné)
        position = {
          left: rect.left + (rect.width / 2),
          top: rect.top + window.scrollY - 5
        };
      }
      
      // Afficher le popup
      setFormatPopup({
        visible: true,
        position
      });
    }
  }, []);

  // Mettre à jour le contenu après formatage
  const updateFormattedContent = useCallback((element, elementType, slideIndex, elementIndex, rowIndex, cellIndex) => {
    // Récupérer le nouveau contenu HTML avec formatage
    const updatedContent = element.innerHTML;
    
    const currentPresentation = presentations[currentPresIndex];
    if (!currentPresentation) return;
    
    const updatedPresentations = [...presentations];
    const currentSlides = convertProjectDataToSlides(currentPresentation.project_data);
    const updatedSlides = [...currentSlides];
    
    if (elementType === 'table' && !isNaN(slideIndex) && !isNaN(elementIndex) && 
        !isNaN(rowIndex) && !isNaN(cellIndex)) {
      // Mise à jour d'une cellule de tableau
      const slide = updatedSlides[slideIndex];
      const tableObject = slide.structuredContent[elementIndex]?.tableObject;
      
      if (tableObject && tableObject.rows && tableObject.rows[rowIndex] && tableObject.rows[rowIndex][cellIndex]) {
        // Mettre à jour avec le contenu HTML formaté
        tableObject.rows[rowIndex][cellIndex].text = updatedContent;
        tableObject.rows[rowIndex][cellIndex].html = true;
        
        // Mettre à jour project_data en fonction de la cellule modifiée
        const updatedProjectData = { ...currentPresentation.project_data };
        
        // Si on modifie une information sur un projet (cellule d'information)
        if (rowIndex > 0 && cellIndex === 1) {
          const projectName = tableObject.rows[rowIndex][0].text;
          
          if (updatedProjectData.activities && updatedProjectData.activities[projectName]) {
            updatedProjectData.activities[projectName].information = updatedContent;
            
            // Extraire les alertes du contenu formaté HTML
            updatedProjectData.activities[projectName].alerts = extractAlertsFromHtml(updatedContent);
          }
        }
        // Si on modifie les événements à venir
        else if (cellIndex === 2) {
          updatedProjectData.upcoming_events = updatedContent;
        }
        
        // Mettre à jour la présentation
        updatedPresentations[currentPresIndex] = {
          ...currentPresentation,
          project_data: updatedProjectData
        };
        
        // Mettre à jour l'état
        setPresentations(updatedPresentations);
      }
    } else if (elementType === 'text') {
      // Mise à jour du texte ordinaire
      const slide = updatedSlides[slideIndex];
      if (slide.structuredContent && slide.structuredContent[elementIndex]) {
        // Mettre à jour le contenu
        slide.structuredContent[elementIndex].content = updatedContent;
        slide.structuredContent[elementIndex].html = true;
        
        // Mettre à jour le project_data si nécessaire
        if (slideIndex === 0 && elementIndex === 0) {
          // Mise à jour du titre
          const updatedProjectData = { ...currentPresentation.project_data };
          updatedProjectData.metadata = { ...updatedProjectData.metadata, title: updatedContent };
          
          updatedPresentations[currentPresIndex] = {
            ...currentPresentation,
            project_data: updatedProjectData
          };
          
          setPresentations(updatedPresentations);
        }
      }
    } else if (elementType === 'title') {
      // Mise à jour du titre
      const slide = updatedSlides[slideIndex];
      slide.title = updatedContent;
      
      // Mettre à jour le project_data si nécessaire
      if (slideIndex === 0) {
        const updatedProjectData = { ...currentPresentation.project_data };
        updatedProjectData.metadata = { ...updatedProjectData.metadata, title: updatedContent };
        
        updatedPresentations[currentPresIndex] = {
          ...currentPresentation,
          project_data: updatedProjectData
        };
        
        setPresentations(updatedPresentations);
      }
    }
  }, [presentations, currentPresIndex, setPresentations, extractAlertsFromHtml, convertProjectDataToSlides]);

  // Formater le texte sélectionné
  const formatSelectedText = useCallback((format) => {
    if (!currentSelection) return;
    
    const { element, range } = currentSelection;
    
    if (!element || !element.isContentEditable) {
      closeFormatPopup();
      return;
    }
    
    // Récupérer les attributs data- de l'élément pour identifier le texte modifié
    const slideIndex = parseInt(element.getAttribute('data-slide-index'));
    const elementType = element.getAttribute('data-element-type');
    const elementIndex = parseInt(element.getAttribute('data-element-index'));
    const rowIndex = parseInt(element.getAttribute('data-row-index'));
    const cellIndex = parseInt(element.getAttribute('data-cell-index'));
    
    // Sauvegarder la sélection actuelle
    const selectedText = range.toString();
    
    // Méthode plus fiable pour insérer du HTML formaté
    try {
      // Créer le contenu formaté
      let formattedHtml = '';
      if (format === 'green') {
        formattedHtml = `<span style="color: green;">${selectedText}</span>`;
      } else if (format === 'red') {
        formattedHtml = `<span style="color: red; font-weight: bold;">${selectedText}</span>`;
      } else if (format === 'black') {
        formattedHtml = selectedText; // Texte normal sans formatage
      } else if (format === 'italic') {
        formattedHtml = `<span style="font-style: italic;">${selectedText}</span>`;
      }
      
      // Supprimer le contenu sélectionné
      range.deleteContents();
      
      // Créer un élément temporaire pour contenir le HTML formaté
      const fragment = document.createDocumentFragment();
      const div = document.createElement('div');
      div.innerHTML = formattedHtml;
      
      // Ajouter tous les nœuds du div au fragment
      while (div.firstChild) {
        fragment.appendChild(div.firstChild);
      }
      
      // Insérer le fragment dans la selection
      range.insertNode(fragment);
      
      // Réinitialiser la sélection
      window.getSelection().removeAllRanges();
      
      // Mettre à jour le contenu dans les données
      setTimeout(() => {
        updateFormattedContent(element, elementType, slideIndex, elementIndex, rowIndex, cellIndex);
      }, 50);
    } catch (error) {
      console.error('Erreur lors du formatage du texte:', error);
    }
    
    // Fermer le popup
    closeFormatPopup();
  }, [currentSelection, closeFormatPopup, updateFormattedContent]);

  // Handle content edit in the structured preview - avec useCallback
  const handleStructuredContentEdit = useCallback((event) => {
    const element = event.target;
    const slideIndex = parseInt(element.getAttribute('data-slide-index'));
    const elementType = element.getAttribute('data-element-type');
    const updatedContent = element.innerText;
    
    if (isNaN(slideIndex) || !elementType) return;
    
    const currentPresentation = presentations[currentPresIndex];
    if (!currentPresentation) return;
    
    const updatedPresentations = [...presentations];
    const currentSlides = convertProjectDataToSlides(currentPresentation.project_data);
    const updatedSlides = [...currentSlides];
    const slide = updatedSlides[slideIndex];
    
    // Handle title edit
    if (elementType === 'title') {
      // Update slide title
      slide.title = updatedContent;
      
      // Update metadata title if editing the main slide
      if (slideIndex === 0) {
        const updatedProjectData = { ...currentPresentation.project_data };
        updatedProjectData.metadata = { ...updatedProjectData.metadata, title: updatedContent };
        
        // Update the presentation data
        updatedPresentations[currentPresIndex] = {
          ...currentPresentation,
          project_data: updatedProjectData
        };
      }
    } 
    // Handle text edit
    else if (elementType === 'text') {
      // Update text in structured content
      const elementIndex = parseInt(element.getAttribute('data-element-index'));
      if (!isNaN(elementIndex) && slide.structuredContent && slide.structuredContent[elementIndex]) {
        slide.structuredContent[elementIndex].content = updatedContent;
      }
    }
    // Handle table cell edit
    else if (elementType.includes('table')) {
      // Update cell in table
      const elementIndex = parseInt(element.getAttribute('data-element-index'));
      const rowIndex = parseInt(element.getAttribute('data-row-index'));
      const cellIndex = parseInt(element.getAttribute('data-cell-index'));
      
      // Get the table from structured content
      if (!isNaN(elementIndex) && slide.structuredContent && slide.structuredContent[elementIndex] && 
          slide.structuredContent[elementIndex].tableObject) {
        
        const tableObject = slide.structuredContent[elementIndex].tableObject;
        
        // Check if row and cell exist
        if (tableObject.rows && tableObject.rows[rowIndex] && tableObject.rows[rowIndex][cellIndex]) {
          // Update the cell text
          tableObject.rows[rowIndex][cellIndex].text = updatedContent;
          
          // Update the presentation data based on cell position
          const updatedProjectData = { ...currentPresentation.project_data };
          
          // If editing header row (row 0), nothing to update in project data
          if (rowIndex === 0) {
            // Nothing to do for header row
          }
          // If editing a project name (first column of data rows)
          else if (cellIndex === 0) {
            // Get the original project name
            const originalProjectName = tableObject.rows[rowIndex][cellIndex].text;
            
            // Update project name if it exists in activities
            if (updatedProjectData.activities && updatedProjectData.activities[originalProjectName]) {
              // Store the project info
              const projectInfo = updatedProjectData.activities[originalProjectName];
              
              // Delete the old key and create a new one with the updated name
              delete updatedProjectData.activities[originalProjectName];
              updatedProjectData.activities[updatedContent] = projectInfo;
            }
          }
          // If editing project information (second column of data rows)
          else if (cellIndex === 1) {
            // Get the project name from the first cell in this row
            const projectName = tableObject.rows[rowIndex][0].text;
            
            // Update project information if the project exists
            if (updatedProjectData.activities && updatedProjectData.activities[projectName]) {
              updatedProjectData.activities[projectName].information = updatedContent;
              
              // Clear all alerts (they will be re-detected if present in the edited text)
              updatedProjectData.activities[projectName].alerts = {
                advancements: [],
                small_alerts: [],
                critical_alerts: []
              };
              
              // Re-detect color-coded alerts if they exist in the updated text
              // This would need more sophisticated parsing to preserve color formatting
              // For now, we'll rely on the updateProjectData function to re-extract them
            }
          }
          // If editing upcoming events (third column)
          else if (cellIndex === 2) {
            updatedProjectData.upcoming_events = updatedContent;
          }
          
          // Update the presentation data
          updatedPresentations[currentPresIndex] = {
            ...currentPresentation,
            project_data: updatedProjectData
          };
        }
      }
    }
    
    // Update state
    setPresentations(updatedPresentations);
  }, [presentations, currentPresIndex]);

  // Add event listeners for editable content after the structured preview is rendered
  useEffect(() => {
    const structuredPreview = document.querySelector('.slide-structured-preview');
    if (structuredPreview) {
      // Remove existing listeners first to prevent duplicates
      const editableElements = structuredPreview.querySelectorAll('.editable');
      editableElements.forEach(element => {
        element.removeEventListener('blur', handleStructuredContentEdit);
        element.removeEventListener('mouseup', handleTextSelection);
      });
      
      // Add listeners to new elements
      editableElements.forEach(element => {
        element.addEventListener('blur', handleStructuredContentEdit);
        element.addEventListener('mouseup', handleTextSelection);
      });
      
      // Add document listener for closing the popup
      document.addEventListener('mousedown', handleClickOutside);
      
      // Cleanup when component unmounts
      return () => {
        editableElements.forEach(element => {
          element.removeEventListener('blur', handleStructuredContentEdit);
          element.removeEventListener('mouseup', handleTextSelection);
        });
        document.removeEventListener('mousedown', handleClickOutside);
      };
    }
  }, [currentPresIndex, selectedSlide, presentations, formatPopup.visible, handleStructuredContentEdit, handleTextSelection, handleClickOutside]);

  // Submit folder name - conserver cette fonction pour permettre aussi la saisie manuelle
  const handleFolderSubmit = async (e) => {
    e.preventDefault();
    let folderToUse = folderName;
    
    // Si on n'a pas de nom de dossier mais qu'on a un chatId, utiliser le chatId
    if (!folderToUse && chatId) {
      folderToUse = chatId;
      setFolderName(chatId);
    }
    
    // Ne rien faire s'il n'y a pas de dossier/chatId
    if (!folderToUse) {
      setError('Veuillez saisir un nom de dossier ou sélectionner une conversation');
      return;
    }
    
    await fetchPresentations(folderToUse);
  };
  
  // Update project data when slides change
  const updateProjectData = (slides, originalData) => {
    if (!slides || slides.length === 0) return originalData;
    
    const projectData = {
      activities: {},
      metadata: { ...originalData.metadata }
    };
    
    // Process the main slide (first slide)
    const mainSlide = slides[0];
    
    // Extract upcoming events from the main slide's table
    // It should be in the third column of the first row
    if (mainSlide.tables && mainSlide.tables[0] && mainSlide.tables[0].rows && mainSlide.tables[0].rows.length > 1) {
      const table = mainSlide.tables[0];
      const firstDataRow = table.rows[1]; // Index 1 is first data row after header
      
      if (firstDataRow.length >= 3) {
        const upcomingEvents = firstDataRow[2].text || '';
        projectData.upcoming_events = upcomingEvents;
      }
    }
    
    // Process table rows to extract project information
    if (mainSlide.tables && mainSlide.tables[0] && mainSlide.tables[0].rows) {
      const table = mainSlide.tables[0];
      
      // Skip header row (index 0)
      for (let i = 1; i < table.rows.length; i++) {
        const row = table.rows[i];
        if (row.length < 2) continue;
        
        const projectName = row[0].text || '';
        if (!projectName) continue;
        
        const information = row[1].text || '';
        
        // Initialize project
        projectData.activities[projectName] = {
          information,
          alerts: {
            advancements: [],
            small_alerts: [],
            critical_alerts: []
          }
        };
        
        // Extract alerts by analyzing the HTML content
        const htmlContent = row[1].text || '';
        
        // Find advancements (green text)
        const advancementMatches = htmlContent.match(/<span style="color: green;">(.*?)<\/span>/g);
        if (advancementMatches) {
          advancementMatches.forEach(match => {
            const content = match.replace(/<span style="color: green;">/, '').replace(/<\/span>/, '');
            projectData.activities[projectName].alerts.advancements.push(content);
          });
        }
        
        // Find small alerts (orange text)
        const smallAlertMatches = htmlContent.match(/<span style="color: orange;">(.*?)<\/span>/g);
        if (smallAlertMatches) {
          smallAlertMatches.forEach(match => {
            const content = match.replace(/<span style="color: orange;">/, '').replace(/<\/span>/, '');
            projectData.activities[projectName].alerts.small_alerts.push(content);
          });
        }
        
        // Find critical alerts (red text)
        const criticalAlertMatches = htmlContent.match(/<span style="color: red; font-weight: bold;">(.*?)<\/span>/g);
        if (criticalAlertMatches) {
          criticalAlertMatches.forEach(match => {
            const content = match.replace(/<span style="color: red; font-weight: bold;">/, '').replace(/<\/span>/, '');
            projectData.activities[projectName].alerts.critical_alerts.push(content);
          });
        }
      }
    }
    
    return projectData;
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
          // Skip title if already shown
          if (item.isTitle) return;
          
          // Utiliser le contenu comme HTML si html=true est défini
          const content = item.html ? item.content : escapeHtml(item.content);
          
          previewContent += `<div 
            class="preview-text editable" 
            contenteditable="true"
            data-slide-index="${slideIndex}" 
            data-element-type="text"
            data-element-index="${itemIndex}"
          >${content}</div>`;
        } else if (item.type === 'table') {
          previewContent += `<div class="table-container">`;
          previewContent += `<table class="preview-table">`;
          
          // Add table with proper structure and editable cells
          const tableObj = item.tableObject || { rows: item.content.map(row => row.map(cell => ({ text: cell }))) };
          
          // Add column group if we have widths
          if (tableObj.colWidths && tableObj.colWidths.length > 0) {
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
            headerRow.forEach((cell, cellIdx) => {
              const style = cell.style ? 
                `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
              
              // Utiliser le contenu comme HTML si html=true est défini
              const cellContent = cell.html ? cell.text : escapeHtml(cell.text || '');
              
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
                data-cell-index="${cellIdx}"
              >${cellContent}</th>`;
            });
            
            previewContent += `</tr></thead><tbody>`;
            
            // Render body rows starting from row 1
            tableObj.rows.slice(1).forEach((row, rowIndexOffset) => {
              const rowIndex = rowIndexOffset + 1; // actual index in the rows array
              previewContent += `<tr>`;
              
              row.forEach((cell, cellIdx) => {
                const style = cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                // Utiliser le contenu comme HTML si html=true est défini
                const cellContent = cell.html ? cell.text : escapeHtml(cell.text || '');
                
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
                  data-cell-index="${cellIdx}"
                >${cellContent}</td>`;
              });
              
              previewContent += `</tr>`;
            });
            
            previewContent += `</tbody>`;
          } else {
            // No specific header, render all rows normally
            previewContent += `<tbody>`;
            
            tableObj.rows.forEach((row, rowIndex) => {
              previewContent += `<tr>`;
              
              row.forEach((cell, cellIdx) => {
                const style = cell.style ? 
                  `style="${cell.style.bold ? 'font-weight:bold;' : ''}${cell.style.italic ? 'font-style:italic;' : ''}${cell.style.underline ? 'text-decoration:underline;' : ''}${cell.style.bgColor ? `background-color:#${cell.style.bgColor};` : ''}"` : '';
                
                // Utiliser le contenu comme HTML si html=true est défini
                const cellContent = cell.html ? cell.text : escapeHtml(cell.text || '');
                
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
                  data-cell-index="${cellIdx}"
                >${cellContent}</td>`;
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
    }
    
    previewContent += `</div>`;
    return previewContent;
  };

  // Fonction utilitaire pour échapper le HTML
  const escapeHtml = (unsafe) => {
    if (typeof unsafe !== 'string') return '';
    return unsafe
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
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
    } else {
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
    }
    
    tableHtml += `</table>`;
    return tableHtml;
  };

  // Export PPTX
  const exportPptx = () => {
    try {
      const pptx = new pptxgen();
      const currentPresentation = presentations[currentPresIndex];
      if (!currentPresentation) return;
      
      const slides = convertProjectDataToSlides(currentPresentation.project_data);
      
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
        if (slide.content) {
          pptxSlide.addText(slide.content, {
            x: 1,
            y: 1.5,
            fontSize: 14
          });
        }
        
        // Add tables if present
        if (slide.tables && slide.tables.length > 0) {
          slide.tables.forEach((table, index) => {
            if (table.rows) {
              const tableData = table.rows.map(row => 
                row.map(cell => typeof cell === 'object' ? cell.text : cell)
              );
              
              pptxSlide.addTable(tableData, {
                x: 1,
                y: 3 + (index * 2),
                w: 8,
                border: { pt: 1, color: '666666' },
                fontFace: 'Arial',
                fontSize: 12,
                rowH: 0.5
              });
            }
          });
        }
      });
      
      // Save the PPTX file
      const exportFileName = currentPresentation.filename || 'presentation.pptx';
      pptx.writeFile({ fileName: exportFileName });
    } catch (error) {
      console.error('Error exporting PPTX:', error);
      alert('Failed to export the presentation');
    }
  };

  const renderCurrentPresentation = () => {
    const currentPresentation = presentations[currentPresIndex];
    if (!currentPresentation) return null;
    
    const slides = convertProjectDataToSlides(currentPresentation.project_data);
    
  return (
      <div className="editor-container">
        <div className="editor-header">
          <h2>
            {currentPresentation.filename || 'Présentation sans titre'}
            <span className="presentation-counter">
              {` (${currentPresIndex + 1}/${presentations.length})`}
            </span>
          </h2>
          <div className="editor-actions">
            {presentations.length > 1 && (
              <>
                <button 
                  onClick={() => setCurrentPresIndex(Math.max(0, currentPresIndex - 1))}
                  disabled={currentPresIndex === 0}
                >
                  Présentation précédente
                </button>
                <button 
                  onClick={() => setCurrentPresIndex(Math.min(presentations.length - 1, currentPresIndex + 1))}
                  disabled={currentPresIndex === presentations.length - 1}
                >
                  Présentation suivante
                </button>
              </>
            )}
            <button onClick={exportPptx}>Exporter en PPTX</button>
            <button onClick={() => { setPresentations([]); setFolderName(''); }}>
              Charger un autre dossier
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
                <div className="slide-title-preview">{slide.title}</div>
              </div>
            ))}
      </div>

          <div className="slide-editor">
            {isLoading ? (
              <div className="loading">Chargement de la présentation...</div>
            ) : (
              <>
                <div className="slide-structure-container">
                  <h3>Aperçu structuré <span className="edit-hint">(cliquez sur les éléments pour les modifier)</span></h3>
                  <div className="structured-preview">
                    <div 
                      dangerouslySetInnerHTML={{ 
                        __html: generateEditableSlidePreview(slides[selectedSlide], selectedSlide) 
                      }} 
                      className="slide-structured-preview"
                    />
                    {formatPopup.visible && (
                      <TextFormatPopup 
                        position={formatPopup.position}
                        onFormat={formatSelectedText}
                        onClose={closeFormatPopup}
                      />
                    )}
                  </div>
                </div>
                
              </>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="pptx-viewer-container">
      <div className="viewer-header">
        <h2>Visualiseur de présentations</h2>
        
        {/* Formulaire de saisie manuelle du Chat ID */}
        <div className="manual-chatid-form">
          <form onSubmit={(e) => {
            e.preventDefault();
            const input = e.target.elements.chatIdInput;
            if (input && input.value) {
              console.log("Utilisation manuelle du Chat ID:", input.value);
              setFolderName(input.value);
              fetchPresentations(input.value);
              input.value = '';
            }
          }}>
            <div className="input-group">
              <input 
                type="text" 
                name="chatIdInput" 
                placeholder="Entrez un Chat ID manuellement" 
                className="chatid-input"
              />
              <button type="submit" className="chatid-submit">Charger</button>
            </div>
            <small className="help-text">
              {chatId ? `Chat ID actuel: ${chatId}` : "Aucun Chat ID sélectionné"}
            </small>
          </form>
        </div>
        
        <div className="presentation-controls">
          {/* Navigation entre présentations */}
        </div>
      </div>
      
      {presentations.length === 0 ? (
        <div className="folder-input-container">
          <h2>Analyser les présentations</h2>
          {chatId ? (
            <div className="chat-info">
              <p>Conversation sélectionnée: <strong>{chatId}</strong></p>
              {isLoading ? (
                <p>Chargement des présentations...</p>
              ) : (
                <button 
                  onClick={() => fetchPresentations(chatId)}
                  className="submit-btn"
                >
                  Charger les présentations
                </button>
              )}
            </div>
          ) : (
            <form onSubmit={handleFolderSubmit} className="folder-form">
              <input
                type="text"
                value={folderName}
                onChange={(e) => setFolderName(e.target.value)}
                placeholder="Nom du dossier à analyser"
                className="folder-input"
              />
              <button type="submit" className="submit-btn" disabled={isLoading}>
                {isLoading ? 'Chargement...' : 'Analyser'}
              </button>
            </form>
          )}
          {error && <div className="error-message">{error}</div>}
        </div>
      ) : (
        renderCurrentPresentation()
      )}
    </div>
  );
};

export default PptxViewer;