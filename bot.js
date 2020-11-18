const moment = require('moment-timezone');
const bent = require('bent')
const post = bent('POST', 'json', 200);

const { ActivityHandler, CardFactory, ActionTypes, ActivityTypes, TeamsInfo } = require('botbuilder');

var _users = [];
let _emulatorMemberCount = 0;
const _emulatorMembers = [
    // emulate the 3 different cases of teams info I've seen
    { name: "First Last", givenName: "First", surname: "Last", email: "firstlast@outlook.com" },
    { name: "NameOnly", email: "name@outlook.com" },
    { name: "", email: "noname@outlook.com" }
];
const _nameHacks = {
    // map of emails in Teams to user context object that matches TRIRIGA info
    // ideally, our emails in Teams would match TRIRIGA and then just send that to TRIRIGA Assistant
    //'email Teams has': { name: { first: "First name TRIRIGA has", last: "Last name TRIRIGA has" }, email: 'Email TRIRIGA has' },
    'firstlast@outlook.com': { name: { first: "First", last: "Last" }, email: 'firstlast@outlook.com' },
    'name@outlook.com': { name: { first: "Name", last: "" }, email: "name@outlook.com" },
    'noname@outlook.com': { name: { first: "", last: "" }, email: "noname@outlook.com" }
};

async function callProxy(message, turnContext, tririgaResult) {

    const userId = turnContext.activity.from.id;

    var payload = {
        'sessionId': _users[userId].sessionId,
        'integration_id': process.env.TA_INTEGRATION_ID,
        'wa_payload': {
            'input': {
                'message_type': 'text',
                'text': message,
                'options': {
                    'return_context': true,
                    'debug': true
                }
            },
            'context': {
                "skills": {
                    "main skill": {
                        "user_defined": {
                            "userContext": undefined
                        }
                    }
                }
            }
        }
    };

    // add user context from storage
    payload.wa_payload.context.skills["main skill"].user_defined.userContext = _users[userId].userContext;

    // add results from cloud function if provided
    if (tririgaResult) {
        payload.wa_payload.context.skills["main skill"].user_defined.tririgaResult = tririgaResult;
    }

    let response = {};
    try {
        console.log(">>>> Sending to Proxy:", JSON.stringify(payload));
        response = await post(process.env.TA_PROXY, payload);
        console.log(">>>> Response from Proxy:", JSON.stringify(response));
    } catch (e) {
        console.error("Got error when calling proxy", e);
        throw Error(`Proxy activation ID: ${e.headers["x-openwhisk-activation-id"]}`);
    }

    // if the session has timed out, resend
    if (response.message && response.message === "Invalid Session") {
        _users[userId].sessionId = "";
        // if it was for sure a reply, then have to start over from beginning
        if (turnContext.activity.replyToId) {
            await turnContext.sendActivity("Our last conversation session has timed out.  We will have to start over.");
            await callProxy("hi", turnContext);
        } else {
            // there is a chance that it wasn't a reply, so give it a try
            await turnContext.sendActivity("Our last conversation session has timed out.  Resending request...");
            await callProxy(message, turnContext);
        }
        return;
    }

    // check if they said hi, if didn't and new session, then we will send again
    let saidHi = false;
    const msgLowerCase = message.toLowerCase();
    if (msgLowerCase === "hi" || msgLowerCase === "hello") {
        saidHi = true;
    }

    // if no session and didn't say Hi, resend
    if (_users[userId].sessionId === "" && !saidHi) {
        // store session for resend
        _users[userId].sessionId = response.result.sessionId;
        await turnContext.sendActivity("One moment...");
        await callProxy(message, turnContext);
        return;
    }

    // got this far, so must have a valid sesion, store it
    _users[userId].sessionId = response.result.sessionId;

    // if we got a user context, store it
    try {
        let returnedUserContext = response.result.result.context.skills['main skill'].user_defined.userContext;
        if (returnedUserContext) {
            console.log(">>>> Setting the userContext to:", JSON.stringify(returnedUserContext));
            _users[userId].userContext = returnedUserContext;
        }
    } catch (e) {
        // no userContext returned so okay to keep what we have
    }

    // show response
    await handleProxyResponse(response.result.result, turnContext);
}

async function callCloudFunction(webhookUrl, params) {

    console.log(`>>>> Sending to Cloud Function: ${webhookUrl}`, JSON.stringify(params));
    let response = {};
    try {
        response = await post(webhookUrl, params);
    } catch (e) {
        console.error("Got Error when calling cloud function", e);
        throw Error(`CF activation ID: ${e.headers["x-openwhisk-activation-id"]}`);
    }
    console.log(">>>> Response from Cloud Function:", JSON.stringify(response));
    return response;
}

async function handleProxyResponse(response, turnContext) {

    for (let i = 0; i < response.output.generic.length; i++) {
        let response_type = response.output.generic[i].response_type;
        if (response_type === "text") {
            const text = response.output.generic[i].text;
            const httpLoc = text.search('http');
            if (httpLoc > 0) {
                // found URL
                const endHttpLoc = text.substr(httpLoc, text.length).indexOf(" ");
                const url = text.substr(httpLoc, endHttpLoc).trim();
                const angleRight = text.indexOf(">");
                const angleLeft = text.substr(angleRight, text.length).indexOf("<");
                const anchorText = text.substr(angleRight + 1, angleLeft - 1);
                const reply = { type: ActivityTypes.Message };
                let buttons = [{
                    type: ActionTypes.OpenUrl,
                    title: anchorText,
                    value: url
                }];
                const card = CardFactory.heroCard('', text, buttons);
                reply.attachments = [card];
                await turnContext.sendActivity(reply);
            } else {
                await turnContext.sendActivity(response.output.generic[i].text);
            }
        } else if (response_type === "image") {
            const imageUrl = response.output.generic[i].source;
            const cardTitle = response.output.generic[i].title;
            const cardText = response.output.generic[i].description;
            const card = CardFactory.heroCard(
                cardTitle,
                cardText,
                [imageUrl]
            );
            await turnContext.sendActivity({ attachments: [card] });
        } else if (response_type === "option") {
            const reply = { type: ActivityTypes.Message };
            let buttons = [];
            response.output.generic[i].options.forEach(option => {
                buttons.push({
                    type: ActionTypes.MessageBack,
                    title: option.label,
                    text: option.value.input.text,
                    displayText: option.value.input.text
                });
            });
            const card = CardFactory.heroCard('', undefined, buttons);
            reply.attachments = [card];
            let resource = await turnContext.sendActivity(reply);
            let user = await getUser(turnContext);
            user.optionsToDelete.push(resource.id);
        }
    }

    if (response.output.actions && response.output.actions[0].type === "client") {
        let cfResponse = await callCloudFunction(
            response.context.skills["main skill"].user_defined.private.cloudfunctions.webhook,
            response.output.actions[0].parameters
        );
        await callProxy("", turnContext, cfResponse);
    }
}

async function getUser(turnContext) {

    let userId = turnContext.activity.from.id;
    if (userId in _users) {
        return _users[userId];
    } else {
        let member = {};
        try {
            if (turnContext.activity.channelId === "emulator" || turnContext.activity.channelId === "webchat") {
                member = _emulatorMembers[_emulatorMemberCount++ % _emulatorMembers.length];
            } else {
                member = await TeamsInfo.getMember(turnContext, userId);
                console.log("found member", member);
            }
        } catch (e) {
            console.error(`Error when looking up member with ID ${userId}`, e);
            throw e;
        }
        _users.push(userId);
        _users[userId] = {};
        //_users[userId].userContext = _nameHacks[member.email.toLowerCase()];
        _users[userId].userContext = { email: member.email }
        _users[userId].sessionId = "";
        _users[userId].optionsToDelete = [];
    }
    return _users[userId];
}

class TririgaBot extends ActivityHandler {

    constructor() {

        super();
        this.onMessage(async (turnContext, next) => {
            console.log(">>>> Received from Teams:", JSON.stringify(turnContext.activity));

            let user = await getUser(turnContext);
            console.log("optionsToDelete is", user.optionsToDelete);
            user.optionsToDelete.forEach(async activity => {
                await turnContext.deleteActivity(activity);
            });
            user.optionsToDelete = [];

            if (turnContext.activity.text) {
                await callProxy(turnContext.activity.text, turnContext);
            } else {
                await turnContext.sendActivity("Sorry, I only handle text responses at the moment.");
            }
            await next();
        });
    }
}

module.exports.TririgaBot = TririgaBot;
