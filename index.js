require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);

const sessions = {};

async function getGraphToken() {

  const tokenResponse = await axios.post(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    }),
    {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    }
  );

  return tokenResponse.data.access_token;
}

async function createSubmission(ctx, data) {

  const token = await getGraphToken();

  const payload = {
    fields: {

      Title: `Shift submission - ${ctx.from.first_name || 'Rep'}`,

      RepName: ctx.from.first_name || '',

      TelegramUserID: String(ctx.from.id),

      _x0024_10Donations: data.d10,
      _x0024_20Donations: data.d20,
      _x0024_25Donations: data.d25,
      _x0024_30Donations: data.d30,
      _x0024_35Donations: data.d35,
      _x0024_40Donations: data.d40,

      Status: 'Submitted'

    }
  };

  await axios.post(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items`,
    payload,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
}

bot.start((ctx) => {

  ctx.reply(
    `Welcome ${ctx.from.first_name || 'there'}!\n\nUse /submit to submit your shift report.`
  );

});

bot.command('submit', (ctx) => {

  sessions[ctx.from.id] = {
    step: 'd10',
    data: {}
  };

  ctx.reply('How many $10 donations did you get?');

});

bot.on('text', async (ctx) => {

  const userId = ctx.from.id;

  const text = ctx.message.text.trim();

  if (text.startsWith('/')) return;

  const session = sessions[userId];

  if (!session) return;

  if (text.toLowerCase() === 'cancel') {

    delete sessions[userId];

    return ctx.reply('❌ Submission cancelled.');

  }

  const askNumber = (value) => {

    const num = Number(value);

    return Number.isInteger(num) && num >= 0
      ? num
      : null;

  };

  try {

    switch (session.step) {

      case 'd10':

        session.data.d10 = askNumber(text);

        if (session.data.d10 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'd20';

        return ctx.reply('How many $20 donations did you get?');

      case 'd20':

        session.data.d20 = askNumber(text);

        if (session.data.d20 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'd25';

        return ctx.reply('How many $25 donations did you get?');

      case 'd25':

        session.data.d25 = askNumber(text);

        if (session.data.d25 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'd30';

        return ctx.reply('How many $30 donations did you get?');

      case 'd30':

        session.data.d30 = askNumber(text);

        if (session.data.d30 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'd35';

        return ctx.reply('How many $35 donations did you get?');

      case 'd35':

        session.data.d35 = askNumber(text);

        if (session.data.d35 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'd40';

        return ctx.reply('How many $40 donations did you get?');

      case 'd40':

        session.data.d40 = askNumber(text);

        if (session.data.d40 === null) {
          return ctx.reply('Please enter a valid number.');
        }

        session.step = 'confirm';

        return ctx.reply(
`Confirm submission:

$10: ${session.data.d10}
$20: ${session.data.d20}
$25: ${session.data.d25}
$30: ${session.data.d30}
$35: ${session.data.d35}
$40: ${session.data.d40}

Type YES to submit or CANCEL to cancel.`
        );

      case 'confirm':

        if (
          text.toLowerCase() !== 'yes'
        ) {

          return ctx.reply(
            'Type YES to submit or CANCEL to cancel.'
          );

        }

        await createSubmission(
          ctx,
          session.data
        );

        delete sessions[userId];

        return ctx.reply(
          '✅ End-of-shift submission saved successfully.'
        );

    }

  } catch (error) {

    console.log(
      error.response?.data || error.message
    );

    return ctx.reply(
      '❌ Submission failed. Please contact management.'
    );

  }

});

bot.launch();

console.log(
  'TechSid Telegram Bot is running...'
);

process.once(
  'SIGINT',
  () => bot.stop('SIGINT')
);

process.once(
  'SIGTERM',
  () => bot.stop('SIGTERM')
);
