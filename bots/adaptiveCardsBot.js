// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory, TurnContext, TeamsInfo } = require('botbuilder');
const { AdaptiveCard, TextBlock, HostConfig } = require("adaptivecards");
const { Template } = require("adaptivecards-templating");

// Import Template & Data
const templateJson = require('../resources/CardTemplate.json');
const dataJson = require('../resources/CardData.json');
// Import AdaptiveCard content.
const FlightItineraryCard = require('../resources/FlightItineraryCard.json');
const ImageGalleryCard = require('../resources/ImageGalleryCard.json');
const LargeWeatherCard = require('../resources/LargeWeatherCard.json');
const RestaurantCard = require('../resources/RestaurantCard.json');
const SolitaireCard = require('../resources/SolitaireCard.json');

// Create array of AdaptiveCard content, this will be used to send a random card to the user.
const CARDS = [
    FlightItineraryCard,
    ImageGalleryCard,
    LargeWeatherCard,
    RestaurantCard,
    SolitaireCard
];

const WELCOME_TEXT = 'This bot will introduce you to Adaptive Cards. Type anything to see an Adaptive Card.';

class AdaptiveCardsBot extends ActivityHandler {
    constructor() {
        super();
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to Adaptive Cards Bot  ${ membersAdded[cnt].name }. ${ WELCOME_TEXT }`);
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            // const member = await TeamsInfo.getMember(context);
            // const members = await TeamsInfo.getMembers(context);
            TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();
            if (text.includes('random')) {
                await this.randomActivityAsync(context);
            } else if (text.includes('template')) {
                await this.templateActivityAsync(context);
            } else if (text.includes('sdk')) {
                await this.sdkActivityAsync(context);
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
    
    async randomActivityAsync(context) {
        const randomlySelectedCard = CARDS[Math.floor((Math.random() * CARDS.length - 1) + 1)];
        await context.sendActivity({
            text: 'Here is an Adaptive Card:',
            attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
        });
    }

    async templateActivityAsync(context) {
        // Templating SDK
        // https://docs.microsoft.com/en-us/adaptive-cards/templating/sdk
        // Samples: https://www.adaptivecards.io/samples/
        const templatePayload = templateJson;
        const template = new Template(templatePayload);

        const evaluationContext = dataJson;
        const cardPayload = template.expand(evaluationContext);
        let card = new AdaptiveCard();
        card.parse(cardPayload);

        let jsonCard = card.toJSON();
        await context.sendActivity({
            text: 'Here is an Adaptive Card based on a Template:',
            attachments: [CardFactory.adaptiveCard(jsonCard)]
        });
    }

    async sdkActivityAsync(context) {
        // Adaptive Cards SDK
        // https://docs.microsoft.com/en-us/adaptive-cards/sdk/rendering-cards/javascript/getting-started
        let card = new AdaptiveCard();
        
        let textBlock = new TextBlock();
        textBlock.setText("Hello World");    
        card.addItem(textBlock);

        card.onExecuteAction = function(action) { alert("Ow!"); }

        // Set its hostConfig property unless you want to use the default Host Config
        // Host Config defines the style and behavior of a card
        card.hostConfig = new HostConfig({
            fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
            // More host config options
        });
        let jsonCard = card.toJSON();
        await context.sendActivity({
            text: 'Here is an SDK generated Adaptive Card:',
            attachments: [CardFactory.adaptiveCard(jsonCard)]
        });
    }
}

module.exports.AdaptiveCardsBot = AdaptiveCardsBot;
