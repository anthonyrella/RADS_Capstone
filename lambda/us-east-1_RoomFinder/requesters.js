/**
 * @file 
 * 
 * Contains support methods to communicate with Microsofts Graph API
 * 
 * Methods include:
 * 1) Retrieving Calendars
 * 2) Finding available times
 * 3) Booking Room
 */

'use strict';

// promises library
const Q = require('q');
// https requests library
const request = require('request');

// object to be exported by this module
const requesters = {}; 


/**
 * requesters.retrieveCalendars
 * 
 * Retrieves all the calendars associated with users mailbox
 *
 * @param  {string} token The JWT access token provided by the Alexa Skill
 * @return {promise}      Promise resolved to JSON containing all calendars.
 */
requesters.retrieveCalendars = function retrieveCalendars(token) {
    const deferred = Q.defer();

    const toGet = {
        /* In order to obtain owner, which I require for consistency, the beta endpoint must be used.
         * TODO: When stable versions are updated, change this endpoint. */
        url: 'https://graph.microsoft.com/beta/Users/Me/Calendars',
        headers: {
            authorization: `Bearer ${token}`,
        },
    };


    request.get(toGet, (err, _response, body) => {

      
        const parsedBody = JSON.parse(body);
      

        if (err) {
            deferred.reject(err);
        } else if (parsedBody.error) {
            deferred.reject(parsedBody.error.message);
        } else {
            deferred.resolve(parsedBody.value);
        }
    });
    return deferred.promise;
};

/**
 * requesters.findFreeRoom 
 * 
 * Find a calendar with availability, given the particular calendars to check.
 *
 * @param  {string} token       The JWT access token provided by the Alexa Skill
 * @param  {string} startTime   The start time of the period to check, formatted as ISO-8601 string
 * @param  {string} endTime     The end time of the period to check, formatted as ISO-8601 string
 * @param  {string[]} names     Array containing the names of all calendars to search for
 * @param  {Object} parsedCals  JSON containing all calendars returned by requesters
 * @return {promise}            Promise resolved to JSON object: ownerName, ownerAdress, name
 */
requesters.findFreeRoom = function findFreeRoom(token, startTime, endTime, names, parsedCals) {
    const deferred = Q.defer();
    /* For each calendar:
     * - check if its name is in 'names'.
     * - if it is in 'names', check if it's free.
     * - if it is free, return its owner and name in a JSON.
     *
     * This is done asynchronously to speed up the process. This means a
     * system must be built to register if no calendars were free.
     * TODO: Improve the code that registers that no rooms are free, as it's hacky.*/

    const calendarsTotal = parsedCals.length;
    let calendarsUnavailable = 0;

    parsedCals.forEach((calendar) => {
        if (names.indexOf(calendar.owner.name) >= 0) {
            const calViewUrl = `https://graph.microsoft.com/v1.0/Users/Me/Calendars/${calendar.id.toString()}/calendarView?startDateTime=${startTime}&endDateTime=${endTime}`;

            const toGet = {
                url: calViewUrl,
                headers: {
                    authorization: `Bearer ${token}`,
                },
            };

            request.get(toGet, (err, response, body) => {
                const parsedBody = JSON.parse(body);
                if (err) {
                    deferred.reject(err);
                } else if (parsedBody.error) {
                    deferred.reject(parsedBody.error.message);
                } else if (parsedBody.value && parsedBody.value.length === 0) {
                    deferred.resolve({
                        ownerName: calendar.owner.name,
                        ownerAddress: calendar.owner.address,
                    });
                } else {
                    calendarsUnavailable += 1;
                    if (calendarsUnavailable === calendarsTotal) {
                        deferred.resolve(false);
                    }
                }
            });
        } else {
            calendarsUnavailable += 1;
            if (calendarsUnavailable === calendarsTotal) {
                deferred.resolve(false);
            }
        }
    });
    return deferred.promise;
};

/**
 * requesters.bookRoom - Book a new event on a calendar using parameters as time.
 *
 * @param  {string} token          The JWT access token provided by the Alexa Skill.
 * @param  {string} ownerAddress   Address of owner of calendar to be booked.
 * @param  {string} ownerName      Name of owner of calendar to be booked.
 * @param  {string} startTime      Start time of meeting to post, formatted as ISO-8601 string.
 * @param  {string} endTime        End time of meeting to post, formatted as ISO-8601 string.
 * @return {promise}               Promise resolved to nothing.
 */

requesters.bookRoom = function bookRoom(token, ownerAddress, ownerName, startTime, endTime) {
    const deferred = Q.defer();

    // Event to be made as JSON
    const newEvent = {
        Subject: 'Alexa\'s Meeting',
        Start: {
            DateTime: startTime,
            TimeZone: 'UTC',
        },
        End: {
            DateTime: endTime,
            TimeZone: 'UTC',
        },
        Body: {
            Content: 'This meeting was booked by Alexa.',
            ContentType: 'Text',
        },
        Attendees: [{
            Status: {
                Response: 'NotResponded',
                Time: startTime,
            },
            Type: 'Required',
            EmailAddress: {
                Address: ownerAddress,
                Name: ownerName,
            },
        }],
    };

    const toPost = {
        url: 'https://graph.microsoft.com/v1.0/me/events',
        headers: {
            'content-type': 'application/json',
            authorization: `Bearer ${token}`,
        },
        body: JSON.stringify(newEvent),
    };

    // Posts event

    request.post(toPost, (err, response, body) => {
        // TODO: Parsed body errors due to bad tokens aren't handled properly. Fix needed.
        const parsedBody = JSON.parse(body);

        if (err) {
            deferred.reject(err);
        } else if (parsedBody.error) {
            deferred.reject(parsedBody.error);
        } else {
            deferred.resolve();
        }
    });
    return deferred.promise;
};

// check room schedule
/**
 * requesters.checkRoomAvailability - Check if the room is free for atleast 15 mins all day.
 *
 * @param  {string} token          The JWT access token provided by the Alexa Skill.
 * @param  {string} roomAddress    Address of room of calendar to be check.
 * @param  {string} startTime      Start time of meeting to check, formatted as ISO-8601 string.
 * @param  {string} endTime        End time of meeting to post, formatted as ISO-8601 string.
 * @return {promise}               Promise resolved to nothing.
 */

requesters.CheckRoomAvailability = function checkRoomAvail(token, ownerAddress, startTime, endTime) {
    const deferred = Q.defer();

    // Event to be made as JSON
    const newEvent = {
        schedules: [
            ownerAddress
        ],
        startTime: {
            dateTime: startTime,
            timeZone: 'UTC'
        },
        endTime: {
            dateTime: endTime,
            timeZone: 'UTC'
        },
        availabilityViewInterval: '15'
    }

    const toPost = {
        url: 'https://graph.microsoft.com/beta/me/calendar/getschedule',
        headers: {
            'content-type': 'application/json',
            authorization: `Bearer ${token}`,
        },
        body: JSON.stringify(newEvent),
    };

    // Posts event
    request.post(toPost, (err, response, body) => {
        // TODO: Parsed body errors due to bad tokens aren't handled properly. Fix needed.
        const parsedBody = JSON.parse(body);

        if (err) {
            deferred.reject(err);
        } else if (parsedBody.error) {
            deferred.reject(parsedBody.error);
        } else {
            deferred.resolve({
                // getting basic values to prove proof of concept
                roomAddress: parsedBody.value[0].scheduleId,
                numberOfMeetings: parsedBody.value[0].scheduleItems.length,
                firstMeetingTime: parsedBody.value[0].scheduleItems[0].start.dateTime
            });
        }
    });
    return deferred.promise;
};

// exports the requesters object
module.exports = requesters; 
