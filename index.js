const path = require('path');
const restify = require('restify');

const dotenv = require('dotenv');
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

console.log("app id is", process.env.MicrosoftAppId);
console.log("password is", process.env.MicrosoftAppPassword);

if (!process.env.MicrosoftAppId) {
    throw Error("Missing environment variable: MicrosoftAppId");
}
if (!process.env.MicrosoftAppPassword) {
    throw Error("Missing environment variable: MicrosoftAppPassword");
}
if (!process.env.TA_PROXY) {
    throw Error("Missing environment variable: TA_PROXY");
}
if (!process.env.TA_INTEGRATION_ID) {
    throw Error("Missing environment variable: TA_INTEGRATION_ID");
}

const { BotFrameworkAdapter } = require('botbuilder');

const { TririgaBot } = require('./bot');

const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${ server.name } listening to ${ server.url }`);
});

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    await context.sendActivity('The bot encountered an error or bug.  Please try again.');
};

adapter.onTurnError = onTurnErrorHandler;

const bot = new TririgaBot();

server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        await bot.run(context);
    });
});
