import re
import time
import json

def format_model_response(response_str: str, model: str = "unknown") -> dict:
    """
    Transforme la réponse du modèle en un format JSON structuré compatible avec OpenWebUI.
    
    La réponse attendue peut contenir un bloc <think>...</think> pour le raisonnement.
    Le contenu entre ces balises est extrait dans "reasoning", et le reste devient "content".
    
    Une fois que le modèle ne renvoie plus de contenu dans le bloc <think>,
    le champ "finish_reason" est fixé à "stop".
    
    Exemple de réponse attendue :
    <think>
    Le raisonnement ici
    </think>
    La réponse principale ici

    Retourne :
    {
      "id": "chatcmpl-<timestamp>",
      "object": "chat.completion",
      "created": <timestamp>,
      "model": "<model>",
      "choices": [
        {
          "index": 0,
          "message": {
             "reasoning": "Le raisonnement ici",
             "content": "La réponse principale ici"
          },
          "finish_reason": "stop"  // si le contenu principal existe
        }
      ],
      "usage": {}
    }
    """
    # Recherche du bloc <think>...</think>
    reasoning_match = re.search(r'<think>(.*?)</think>', response_str, re.DOTALL)
    if reasoning_match:
        reasoning = reasoning_match.group(1).strip()
        # Supprime le bloc <think>...</think> pour récupérer le contenu principal
        content = re.sub(r'<think>.*?</think>', '', response_str, flags=re.DOTALL).strip()
    else:
        reasoning = ""
        content = response_str.strip()
    
    # Dès qu'il y a du contenu principal (en dehors des balises <think>), on considère que la génération est terminée
    finish_reason = "stop" if content else None

    return {
        "id": f"chatcmpl-{int(time.time())}",
        "object": "chat.completion",
        "created": int(time.time()),
        "model": model,
        "choices": [
            {
                "index": 0,
                "message": {
                    "reasoning": reasoning,
                    "content": content
                },
                "finish_reason": finish_reason
            }
        ],
        "usage": {}
    }
