import React, { useState, useEffect } from 'react';
import './FileList.css';

function FileList({ openPresentation, openPresentations }) {
  const [files, setFiles] = useState([
    { id: 1, name: 'Presentation1.pptx', date: '2023-05-20' },
    { id: 2, name: 'Presentation2.pptx', date: '2023-05-22' },
    { id: 3, name: 'Rapport Annuel.pptx', date: '2023-06-15' },
    { id: 4, name: 'Présentation Client.pptx', date: '2023-07-10' },
    { id: 5, name: 'Démonstration Produit.pptx', date: '2023-08-05' },
    { id: 6, name: 'Marketing Stratégie.pptx', date: '2023-09-12' },
    { id: 7, name: 'Budget 2024.pptx', date: '2023-10-18' },
    { id: 8, name: 'Analyse Trimestrielle.pptx', date: '2023-11-07' },
    { id: 9, name: 'Rétrospective 2023.pptx', date: '2023-12-15' },
    { id: 10, name: 'Roadmap 2024.pptx', date: '2024-01-05' },
  ]);

  // Écouter les événements d'ajout de fichiers
  useEffect(() => {
    const handleAddToFileList = (event) => {
      const { file } = event.detail;
      
      // Vérifier si le fichier n'est pas déjà dans la liste
      if (!files.some(f => f.id === file.id)) {
        setFiles(prevFiles => [...prevFiles, file]);
      }
    };

    document.addEventListener('add-to-file-list', handleAddToFileList);

    return () => {
      document.removeEventListener('add-to-file-list', handleAddToFileList);
    };
  }, [files]);

  const savePresentation = () => {
    // Récupérer la présentation active
    const activePresentation = openPresentations.length > 0 
      ? openPresentations.find(p => isFileOpen(p.id)) 
      : null;
    
    if (!activePresentation) {
      alert("Aucune présentation active à sauvegarder !");
      return;
    }
    
    // Simuler la sauvegarde
    alert(`Présentation "${activePresentation.name}" sauvegardée !`);
  };

  const openFile = (file) => {
    openPresentation(file);
  };

  const deleteFile = (id, fileName) => {
    if (window.confirm(`Êtes-vous sûr de vouloir supprimer ${fileName}?`)) {
      setFiles(files.filter(file => file.id !== id));
      alert(`Fichier ${fileName} supprimé !`);
    }
  };

  // Vérifier si un fichier est actuellement ouvert
  const isFileOpen = (id) => {
    return openPresentations.some(p => p.id === id);
  };

  return (
    <div className="file-list">
      <h2>EXPLORER</h2>
      <div className="file-actions">
        <button onClick={savePresentation}>Save Current Presentation</button>
      </div>
      <div className="files-container">
        <ul className="files-list">
          {files.map(file => (
            <li 
              key={file.id} 
              className={`file-item ${isFileOpen(file.id) ? 'active' : ''}`}
              onClick={() => openFile(file)}
            >
              <div className="file-icon">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                  <line x1="16" y1="13" x2="8" y2="13"></line>
                  <line x1="16" y1="17" x2="8" y2="17"></line>
                  <polyline points="10 9 9 9 8 9"></polyline>
                </svg>
              </div>
              <div className="file-details">
                <span className="file-name">{file.name}</span>
                <span className="file-date">{file.date}</span>
              </div>
              <div className="file-actions-buttons">
                <button 
                  className="action-button delete-button"
                  onClick={(e) => {
                    e.stopPropagation();
                    deleteFile(file.id, file.name);
                  }}
                >
                  ×
                </button>
              </div>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

export default FileList; 