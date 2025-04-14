import React, { useState, useEffect } from 'react';
import './App.css';
import PptxViewer from './PptxViewer';
import FileList from './FileList';
import ChatBot from './chatBot';

function App() {
  // État pour gérer les présentations ouvertes
  const [openPresentations, setOpenPresentations] = useState([]);
  // État pour suivre la présentation active
  const [activePresentation, setActivePresentation] = useState(null);
  // État pour suivre le chatId
  const [chatId, setChatId] = useState(null);
  
  // Configurer les écouteurs d'événements pour la communication entre composants
  useEffect(() => {
    const handleOpenPresentation = (event) => {
      const { presentation } = event.detail;
      
      // Si le fichier est déjà ouvert, juste le rendre actif
      if (openPresentations.some(p => p.id === presentation.id)) {
        setActivePresentation(presentation.id);
        return;
      }
      
      // Sinon, ajouter à la liste des présentations ouvertes
      setOpenPresentations(prevPresentations => [...prevPresentations, presentation]);
      
      // Définir comme active
      setActivePresentation(presentation.id);
      
      // Si c'est un fichier importé, l'ajouter aussi à la liste des fichiers dans FileList
      if (presentation.file) {
        // Envoyer un événement pour ajouter le fichier à la liste des fichiers
        const addToFileListEvent = new CustomEvent('add-to-file-list', {
          detail: { 
            file: {
              id: presentation.id,
              name: presentation.name,
              date: presentation.date
            } 
          }
        });
        document.dispatchEvent(addToFileListEvent);
      }
    };

    // Ajouter l'écouteur
    document.addEventListener('open-presentation', handleOpenPresentation);

    // Nettoyer l'écouteur quand le composant est démonté
    return () => {
      document.removeEventListener('open-presentation', handleOpenPresentation);
    };
  }, [openPresentations]);

  // Fonction pour ouvrir une présentation
  const openPresentation = (file) => {
    // Vérifier si le fichier est déjà ouvert
    if (!openPresentations.some(p => p.id === file.id)) {
      // Ajouter à la liste des présentations ouvertes
      setOpenPresentations([...openPresentations, file]);
    }
    // Définir comme active
    setActivePresentation(file.id);
  };

  // Fonction pour fermer une présentation
  const closePresentation = (id) => {
    const newOpenPresentations = openPresentations.filter(p => p.id !== id);
    setOpenPresentations(newOpenPresentations);
    
    // Si la présentation fermée était active, sélectionner une autre
    if (activePresentation === id && newOpenPresentations.length > 0) {
      setActivePresentation(newOpenPresentations[newOpenPresentations.length - 1].id);
    } else if (newOpenPresentations.length === 0) {
      setActivePresentation(null);
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>ACRA</h1>
      </header>
      <main className="App-main">
        <div className="grid-container">
          <div className="files-section">
            <FileList 
              openPresentation={openPresentation}
              openPresentations={openPresentations}
            />
          </div>
          <div className="pptx-section">
            <PptxViewer 
              openPresentations={openPresentations}
              activePresentation={activePresentation}
              setActivePresentation={setActivePresentation}
              closePresentation={closePresentation}
              chatId={chatId}
            />
          </div>
          <div className="chatbot-section">
            <ChatBot setChatId={setChatId}/>
          </div>
        </div>
      </main>
    </div>
  );
}

export default App; 