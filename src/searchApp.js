const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");

class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    // const searchQuery = query.parameters[0].value;

    const attachments = [];

    const template = new ACData.Template(helloWorldCard);
    const card = template.expand({
      $root: {
        name: "Your Communication Manifesto",
        description:
          "1. Be kind. 2. Be honest. 3. Be open. 4. Be inclusive. 5. Be respectful. 6. If you criticize, do it constructively. 7. Keep emails short, under 200 words. 8. Use emojis to convey tone.",
      },
    });
    const preview = CardFactory.heroCard("Your Communication Manifesto");
    const attachment = { ...CardFactory.adaptiveCard(card), preview };
    attachments.push(attachment);

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}

module.exports.SearchApp = SearchApp;
