require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);

const sessions = {};

// ======================
// GET MICROSOFT TOKEN
// ======================

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

// ======================
// GET ACTIVE BOT USER
// ======================

async function getBotUser(telegramId) {

  const token = await getGraphToken();

  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items?expand=fields`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  const users = response.data.value;

  const foundUser = users.find(
    user =>
      user.fields.TelegramUserID === String(telegramId) &&
      user.fields.Active === true
  );

  return foundUser || null;

}

// ======================
// CHECK TODAY SUBMISSION
// ======================

async function hasSubmittedToday(telegramId) {

  const token = await getGraphToken();

  const today = new Date();

  const startOfDay = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate()
  ).toISOString();

  const endOfDay = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate() + 1
  ).toISOString();

  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  return response.data.value.some(item =>
    item.fields.TelegramUserID === String(telegramId) &&
    item.fields.ShiftDate >= startOfDay &&
    item.fields.ShiftDate < endOfDay
  );

}

// ======================
// CREATE SUBMISSION
// ======================

async function createSubmission(ctx, data, userData) {

  const token = await getGraphToken();

  const totalDonations =
    data.d10 +
    data.d20 +
    data.d25 +
    data.d30 +
    data.d35 +
    data.d40;

  const payload = {
    fields: {

      Title: `Shift submission - ${userData.RepName}`,

      RepName: userData.RepName,

      RepEmail: userData.RepEmail,

      TelegramUserID: String(ctx.from.id),

      ShiftDate: new Date().toISOString(),

      TLName: userData.TLName,

      Market_x002f_City: userData.MarketCity,

      _x0024_10Donations: data.d10,
      _x0024_20Donations: data.d20,
      _x0024_25Donations: data.d25,
      _x0024_30Donations: data.d30,
      _x0024_35Donations: data.d35,
      _x0024_40Donations: data.d40,

      TotalDonations: totalDonations,

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

// ======================
// START
// ======================

bot.start(async (ctx) => {

  const user = await getBotUser(ctx.from.id);

  if (!user) {

    return ctx.reply(
      `❌ You are not authorized.\n\nPlease contact management and provide this Telegram ID:\n${ctx.from.id}`
    );

  }

  ctx.reply(
    `Welcome ${user.fields.LinkTitle}!\n\nUse /submit to submit your shift report.`
  );

});

// ======================
// HELP
// ======================

bot.command('help', (ctx) => {

  ctx.reply(
`Commands:

/start
/submit
/teamtoday
/help`
  );

});

// ======================
// TEAM TODAY
// ======================

bot.command('teamtoday', async (ctx) => {

  try {

    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

    if (user.fields.Role !== 'TL') {
      return ctx.reply('❌ Only Team Leads can use this command.');
    }

    const token = await getGraphToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const submissions = response.data.value;

    const today = new Date();

    const startOfDay = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    ).toISOString();

    const endOfDay = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate() + 1
    ).toISOString();

    const tlName = user.fields.LinkTitle;

    const todaySubs = submissions.filter(item => {

      return (
        item.fields.TLName === tlName &&
        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
      );

    });

    if (todaySubs.length === 0) {
      return ctx.reply('No team submissions today.');
    }

    let total = 0;

    let message = '📊 Team Today\n\n';

    todaySubs.forEach(item => {

      const rep = item.fields.RepName;

      const repTotal =
        item.fields.TotalDonations || 0;

      total += repTotal;

      message += `${rep} - ${repTotal}\n`;

    });

    message += `\nTeam Total: ${total}`;

    ctx.reply(message);

  } catch (error) {

    console.log(error.response?.data || error.message);

    ctx.reply('❌ Failed to load team report.');

  }

});

// ======================
// SUBMIT
// ======================

bot.command('submit', async (ctx) => {

  const user = await getBotUser(ctx.from.id);

  if (!user) {

    return ctx.reply(
      '❌ You are not authorized to submit.'
    );

  }

  const alreadySubmitted = await hasSubmittedToday(ctx.from.id);

  if (alreadySubmitted) {

    return ctx.reply(
      '❌ You have already submitted your shift report today.'
    );

  }

  sessions[ctx.from.id] = {
    step: 'd10',
    data: {},
    user: {
      RepName: user.fields.LinkTitle || '',
      RepEmail: user.fields.Email || '',
      TLName: user.fields.TL_x002f_MangerName || '',
      MarketCity: user.fields.Market_x002f_City || ''
    }
  };

  ctx.reply('How many $10 donations did you get?');

});

// ======================
// TEXT FLOW
// ======================

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

Total Donations: ${
  session.data.d10 +
  session.data.d20 +
  session.data.d25 +
  session.data.d30 +
  session.data.d35 +
  session.data.d40
}

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
          session.data,
          session.user
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

// ======================
// LAUNCH BOT
// ======================

bot.launch();

console.log(
  'TechSid Telegram Bot is running...'
);

// ======================
// SAFE SHUTDOWN
// ======================

process.once(
  'SIGINT',
  () => bot.stop('SIGINT')
);

process.once(
  'SIGTERM',
  () => bot.stop('SIGTERM')
);
