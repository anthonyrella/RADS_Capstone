/**
 * @file 
 * 
 * Main lambda function file. 
 * 
 * Contains:
 * 1) All intent handlers
 * 2) Support handlers
 * 3) Main export function to lambda
 */

'use strict';

// official node.js alexa v2 sdk.
const Alexa = require('ask-sdk');
// library for date usage
const moment = require('moment-timezone');
// support module to handle the Microsoft Graph API requests
const requesters = require('./requesters');
// support module containing application configuration information
const config = require('./config');

// Launch Handler is called when there is an invocation of skill without a specific intent called
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

// Provides an option for an available room given the parameters. 
const FindRoomHandler = {


    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'FindRoom';
    },
    handle(handlerInput) {

        // Handles eliciting required slots. Delegates the responsibility to the interaction model
        if (handlerInput.requestEnvelope.request.dialogState != "COMPLETED") {
            return handlerInput.responseBuilder
                .addDelegateDirective()
                .getResponse();

        } else {

            const bookingDuration = moment.duration(handlerInput.requestEnvelope.request.intent.slots.Duration.value);
            const startTime = handlerInput.requestEnvelope.request.intent.slots.StartTime.value;
            const endTime = moment(startTime, 'HH:mm').add(bookingDuration.asMinutes(), 'minutes').format('HH:mm');
            const dateOfMeeting = handlerInput.requestEnvelope.request.intent.slots.Date.value;

            // establishes variable to save slot information into the users session
            const attributes = handlerInput.attributesManager.getSessionAttributes();

            // Save dates in attributes as ISO strings, so they can be accessed to post event.
            attributes.startTime = moment.tz(dateOfMeeting + " " + startTime, "America/Toronto").toISOString();
            attributes.endTime = moment.tz(dateOfMeeting + " " + endTime, "America/Toronto").toISOString();
            attributes.duration = bookingDuration.toISOString();
            attributes.durationInMinutes = Math.ceil(parseFloat(bookingDuration.asMinutes()));
            handlerInput.attributesManager.setSessionAttributes(attributes);

            // retrieves all meeting room calendars
            var meetingRoomCalendars = requesters.retrieveCalendars(handlerInput.requestEnvelope.session.user.accessToken)
                .then((parsedCals) => {

                    // locates avaialbe room based on the parameters
                    var freeRoom = requesters.findFreeRoom(
                        handlerInput.requestEnvelope.session.user.accessToken,
                        attributes.startTime,
                        attributes.endTime,
                        config.testNames,
                        parsedCals)
                        .then((creds) => {
                           
                            if(creds) {
                                // if an available room exists, provides the suggestion, with a question to book the room. 
                                const attributes = handlerInput.attributesManager.getSessionAttributes();
                                attributes.ownerAddress = creds.ownerAddress;
                                attributes.ownerName = creds.ownerName;
                                handlerInput.attributesManager.setSessionAttributes(attributes);
                                return handlerInput.responseBuilder
                                    .speak(handlerInput.attributesManager.getSessionAttributes().ownerName + " is available, would you like to book it?")
                                    .withShouldEndSession(false)
                                    .getResponse();
    
                            }else {
                                // provides message if no avaialble room exists. 
                                return handlerInput.responseBuilder
                                    .speak("No rooms are available at this time. Try again with another time")
                                    .withShouldEndSession(true)
                                    .getResponse();
                            }

                        })

                    return freeRoom;

                })

            return meetingRoomCalendars;
        }

    }
};

// TODO: Have the YesIntent invoke the BookHandler with function logic to book room. 
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

// Contains function logic to book room if user responded with Yes. 
// TODO: Pass responsibility to BookHandler to perform book logic. Will need YesIntent for other responsibilities. Make this intent more modular with session variable flags. 
const YesHandler = {

    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'IntentRequest'
            && handlerInput.requestEnvelope.request.intent.name === 'AMAZON.YesIntent';
    },
    handle(handlerInput) {

        // receives current session variables
        const attributes = handlerInput.attributesManager.getSessionAttributes();

        // TODO: Place check here to see if booking room is possible. 
        // Calls bookroom function to perform booking. 
        var meetingRoomCalendars = requesters.bookRoom(
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

        return meetingRoomCalendars;

    }

}

// Handles users no response to booking question. 
// TODO: Make more modular to handler other situations with session varaible check/flag. 
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

// Provides basic skill help to user
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

// Handles users request to stop or cancel interaction. 
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
            .withShouldEndSession(true)
            .getResponse();
    }
};

// Handles ending session.
// TODO: Have this handler called from other handlers when ending session? 
const SessionEndedRequestHandler = {
    canHandle(handlerInput) {
        return handlerInput.requestEnvelope.request.type === 'SessionEndedRequest';
    },
    handle(handlerInput) {
     
        return handlerInput.responseBuilder
        .withShouldEndSession(true)
        .getResponse();
    }
};

// Generic error handler to process skill errors. 
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

// Main export handler. Processes and adds all above handlers. Executes Alexa code. 
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