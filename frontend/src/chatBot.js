import React, { useEffect, useRef, useState } from 'react';

function ChatBot({ setChatId }) {
  const iframeRef = useRef(null);
  const [observer, setObserver] = useState(null);
  
  useEffect(() => {
    // Fonction pour extraire le chat_id à partir de l'URL
    const extractChatIdFromUrl = (url) => {
      try {
        // Format attendu: http://localhost:3030/c/chat_id
        const match = url.match(/\/c\/([^\/\?#]+)/);
        if (match && match[1]) {
          return match[1];
        }
      } catch (error) {
        console.error("Erreur lors de l'extraction du chat_id:", error);
      }
      return null;
    };

    // Fonction pour vérifier si une chaîne ressemble à un UUID
    const isUuid = (str) => {
      const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
      return uuidRegex.test(str);
    };

    // Gestionnaire de message pour l'événement postMessage
    const handleMessage = (event) => {
      // Vérifier l'origine pour la sécurité (accepter localhost:3030)
      if (event.origin !== 'http://localhost:3030' && !event.origin.startsWith('http://localhost:')) {
        return;
      }

      // Vérifier si le message contient du texte et rechercher des UUID
      if (typeof event.data === 'string') {
        // Rechercher des UUID dans le texte du message
        const uuidRegex = /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi;
        const matches = event.data.match(uuidRegex);
        
        if (matches && matches.length > 0) {
          const chatId = matches[0];
          console.log("UUID détecté dans postMessage:", chatId);
          setChatId(chatId);
          return;
        }
        
        // Gérer les notifications au format SSE
        if (event.data.startsWith('data: ')) {
          try {
            // Extraire le JSON de la notification SSE (format: data: {...}\n\n)
            const jsonData = event.data.substring(6).trim();
            const parsedData = JSON.parse(jsonData);
            
            // Vérifier si c'est une notification de changement de chat_id
            if (parsedData && parsedData.type === "chat_id_changed" && parsedData.chatId) {
              console.log("Notification SSE de changement de chat_id reçue:", parsedData.chatId);
              setChatId(parsedData.chatId);
            }
          } catch (error) {
            console.error("Erreur lors du traitement de la notification SSE:", error);
          }
          return;
        }
      }
      
      // Vérifier si c'est un message standard avec chatId
      if (event.data && event.data.chatId) {
        console.log("Chat ID reçu via postMessage:", event.data.chatId);
        setChatId(event.data.chatId);
        return;
      }
      
      // Gérer les notifications spéciales de la pipeline ACRA
      if (event.data && event.data.type === "chat_id_changed") {
        console.log("Notification de changement de chat_id reçue de la pipeline:", event.data.chatId);
        setChatId(event.data.chatId);
      }
    };

    // Surveillance des changements d'URL de l'iframe
    const checkIframeUrl = () => {
      try {
        if (iframeRef.current && iframeRef.current.contentWindow) {
          const currentUrl = iframeRef.current.contentWindow.location.href;
          const chatId = extractChatIdFromUrl(currentUrl);
          
          if (chatId) {
            console.log("Chat ID détecté dans l'URL:", chatId);
            setChatId(chatId);
          }
        }
      } catch (e) {
        // Erreur d'accès cross-origin - c'est normal pour les iframes
        // Cela se produit lorsque l'iframe est sur un domaine différent
        console.debug("Erreur d'accès cross-origin lors de la vérification de l'URL (normal)");
      }
    };

    // Utilisation d'un MutationObserver pour surveiller la console de notre propre page
    const setupConsoleObserver = () => {
      // Observer la console de développement pour capturer les UUIDs écrits
      const originalConsoleLog = console.log;
      console.log = function(...args) {
        // Appeler le console.log original d'abord
        originalConsoleLog.apply(this, args);
        
        // Vérifier chaque argument pour détecter un UUID
        for (const arg of args) {
          if (typeof arg === 'string' && isUuid(arg.trim())) {
            console.debug("UUID détecté dans les logs de console:", arg.trim());
            setChatId(arg.trim());
            break;
          }
        }
      };
      
      return function cleanup() {
        // Restaurer la fonction console.log originale lors du nettoyage
        console.log = originalConsoleLog;
      };
    };

    // Vérifier régulièrement l'URL (toutes les 2 secondes)
    const checkInterval = setInterval(checkIframeUrl, 2000);
    
    // Ajouter l'écouteur d'événements pour les messages
    window.addEventListener('message', handleMessage);
    
    // Configurer l'observateur de console
    const consoleCleanup = setupConsoleObserver();
    
    // Vérifier l'URL lorsque l'iframe est chargé
    if (iframeRef.current) {
      iframeRef.current.addEventListener('load', checkIframeUrl);
    }

    return () => {
      window.removeEventListener('message', handleMessage);
      clearInterval(checkInterval);
      consoleCleanup();
      
      if (iframeRef.current) {
        iframeRef.current.removeEventListener('load', checkIframeUrl);
      }
    };
  }, [setChatId]);

  return (
    <iframe
      ref={iframeRef}
      src="http://localhost:3030"
      style={{ width: '100%', height: '100%', border: 'none', borderRadius: '5px' }}
      title="ChatBot"
      allow="clipboard-write"
    />
  );
}

export default ChatBot;
