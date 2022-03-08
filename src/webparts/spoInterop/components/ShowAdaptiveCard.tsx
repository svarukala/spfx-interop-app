import * as React from 'react';
import styles from './SpoInterop.module.scss';
import { useState, useEffect } from 'react';
import * as AdaptiveCards from "adaptivecards";

function ShowAdaptiveCard (props) {
  const [ssoToken, setSsoToken] = useState<string>();
  const [error, setError] = useState<string>();
  const [resourceOboToken, setResourceOboToken] = useState<string>();
  const [acContent, setAcContent] = useState<JSX.Element>();

  useEffect(() => {

    var card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "size": "Medium",
                "weight": "Bolder",
                "text": "People Picker with Org search enabled"
            },
            {
                "type": "Input.ChoiceSet",
                "choices": [],
                "choices.data": {
                    "type": "Data.Query",
                    "dataset": "graph.microsoft.com/users"
                },
                "id": "people-picker",
                "value": "4cb08dcb-b50e-4ee6-9712-03fd4c746a6c",
                "isMultiSelect": true
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit"
            }
        ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2"
    };
    
    // Build the AdaptiveCard object
    var adaptiveCard = new AdaptiveCards.AdaptiveCard();

    // And configure the HostConfig (font, font size, style, spacing, etc.)
    adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
    fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
    // More host config options
    });

    // Plug into the actions commands execution
    adaptiveCard.onExecuteAction = handleCardActions;

    // Parse the selected card
    adaptiveCard.parse(card);

    // Render the selected card
    var renderedCard = adaptiveCard.render();
    let content = <div ref={(el) => { el && el.appendChild(renderedCard) }} />
    setAcContent(content);
  }, []);
    
  const handleCardActions = (action) => {

    if (action._propertyBag["type"] === "Action.Submit") {
      alert("You pressed a submit button!");
      alert(`Firstname: ${action._processedData.FirstName}\nLastname: ${action._processedData.LastName}\nBirthdate: ${action._processedData.BirthDate}\nFavorite color: ${action._processedData.FavoriteColor}\nDo you like this form: ${action._processedData.DoYouLikeThis}\n`);
    }
    else if (action._propertyBag["type"] === "Action.OpenUrl") {
      window.open(action.url, "_blank");
    }
  }

  return (
      
        <div className={ styles.renderAdaptiveCard }>
            <div className={ styles.container }>
            { acContent }
            </div>
        </div>
      
  );

  
}




export default ShowAdaptiveCard;


/*
<p>You clicked {count} times</p>
<button onClick={() => setCount(count + 1)}>
  Click me
</button>
*/