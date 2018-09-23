'use strict';

const Alexa = require('ask-sdk');
const moment = require('moment');
const requesters = require('./requesters');
const config = require('./config');
const Q = require('q');


const LaunchRequestHandler = {
    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'LaunchRequest';
    },
    handle(handlerInput) {
        const speechText = 'Welcome to Manage Meetings, what would you like to do?';

        return handlerInput.responseBuilder
            .speak(speechText)
            .reprompt(speechText)
            .withSimpleCard('Hello World', speechText)
            .getResponse();
    }
};

const FindRoomHandler = {


    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'FindRoom';
    },
    handle(handlerInput) {

        var deferred = Q.defer();

        const updatedIntent = handlerInput.requestEnvelope.request.intent;

        //if statement to validate duration is appropriate
        if (handlerInput.requestEnvelope.request.dialogState != "COMPLETED") {
            console.log("Request incomplete");
            return handlerInput.responseBuilder
                .addDelegateDirective()
                .getResponse();

        } else {

            const bookingDuration = moment.duration(handlerInput.requestEnvelope.request.intent.slots.Duration.value);
            const startTime = new Date();
            const endTime = new Date(startTime.getTime() + bookingDuration.asMilliseconds());

            const attributes = handlerInput.attributesManager.getSessionAttributes();

            // Save dates in attributes as ISO strings, so they can be accessed to post event.
            attributes.startTime = startTime.toISOString();
            attributes.endTime = endTime.toISOString();
            attributes.duration = bookingDuration.toISOString();
            attributes.durationInMinutes = Math.ceil(parseFloat(bookingDuration.asMinutes()));
            handlerInput.attributesManager.setSessionAttributes(attributes);


            var meetingRoom = requesters.retrieveCalendars(handlerInput.requestEnvelope.session.user.accessToken)
                .then((parsedCals) => {

                    console.log("in step 1");
                    var returno = requesters.findFreeRoom(
                        handlerInput.requestEnvelope.session.user.accessToken,
                        attributes.startTime,
                        attributes.endTime,
                        config.testNames,
                        parsedCals)
                        .then((creds) => {
                            console.log("in step 2");
                            console.log(creds);
                            const attributes = handlerInput.attributesManager.getSessionAttributes();
                            attributes.ownerAddress = creds.ownerAddress;
                            attributes.ownerName = creds.ownerName;
                            handlerInput.attributesManager.setSessionAttributes(attributes);
                            console.log(handlerInput.attributesManager.getSessionAttributes().ownerName);
                            return handlerInput.responseBuilder
                                .speak(handlerInput.attributesManager.getSessionAttributes().ownerName + " is available, would you like to book it?")
                                .withShouldEndSession(false)
                                .getResponse();

                        })

                    return returno;

                })

            return meetingRoom;
        }

    }
};

const BookHandler = {

    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'BookRoom';
    },
    handle(handlerInput) {

        return handlerInput.responseBuilder
            .speak("This is the book handler")
            .getResponse();

    }



};

const YesHandler = {

    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'AMAZON.YesIntent';
    },
    handle(handlerInput) {

        const attributes = handlerInput.attributesManager.getSessionAttributes();

        var meetingRoom =  requesters.bookRoom(
            handlerInput.requestEnvelope.session.user.accessToken,
            attributes.ownerAddress,
            attributes.ownerName,
            attributes.startTime,
            attributes.endTime)
            .then(() => {

                return handlerInput.responseBuilder
                    .speak(attributes.ownerName + " is now Booked!")
                    .reprompt(attributes.ownerName + " is now Booked!")
                    .withSimpleCard(attributes.ownerName + " is now Booked!")
                    .withShouldEndSession(true)
                    .getResponse();

            });
        
            return meetingRoom;

    }

}

const NoHandler = {

    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'AMAZON.NoIntent';
    },
    handle(handlerInput) {
        const speechText = "This is for when saying NO";


        return handlerInput.responseBuilder
            .speak(speechText)
            .reprompt(speechText)
            .withSimpleCard('Say room finder to find and book a room', speechText)
            .getResponse();
    }

}

const HelpIntentHandler = {
    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'AMAZON.HelpIntent';
    },
    handle(handlerInput) {
        const speechText = 'Say room finder to find and book a room';


        return handlerInput.responseBuilder
            .speak(speechText)
            .reprompt(speechText)
            .withSimpleCard('Say room finder to find and book a room', speechText)
            .getResponse();
    }
};

const CancelAndStopIntentHandler = {
    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && (handlerInput.requestEnvelope.request.intent.name === 'AMAZON.CancelIntent'
                || handlerInput.requestEnvelope.request.intent.name === 'AMAZON.StopIntent');
    },
    handle(handlerInput) {
        const speechText = 'Goodbye!';

        return handlerInput.responseBuilder
            .speak(speechText)
            .withSimpleCard('Goodbye!', speechText)
            .getResponse();
    }
};

const SessionEndedRequestHandler = {
    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'SessionEndedRequest';
    },
    handle(handlerInput) {
        //any cleanup logic goes here
        return handlerInput.responseBuilder.getResponse();
    }
};

const ErrorHandler = {
    canHandle() {
        return true;
    },
    handle(handlerInput, error) {
        console.log(`Error handled: ${error.message}`);

        return handlerInput.responseBuilder
            .speak('Sorry, I can\'t understand the command. Please say again.')
            .reprompt('Sorry, I can\'t understand the command. Please say again.')
            .getResponse();
    },
};

exports.handler = Alexa.SkillBuilders.custom()
    .addRequestHandlers(
        LaunchRequestHandler,
        FindRoomHandler,
        BookHandler,
        YesHandler,
        NoHandler,
        HelpIntentHandler,
        CancelAndStopIntentHandler,
        SessionEndedRequestHandler)
    .addErrorHandlers(ErrorHandler)
    .lambda();