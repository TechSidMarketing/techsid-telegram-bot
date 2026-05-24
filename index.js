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
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );
  return tokenResponse.data.access_token;
}

async function createSubmission(ctx, data) {
  const token = await getGraphToken();

  const payload = {
    fields: {
      Title: `Shift submission - ${ctx.from.first_name || 'Rep'}`,
      RepName: ctx.from.first_name || '',
      RepEmail: '',
      TelegramUserID: String(ctx.from.id),
      TLName: '',
      MarketCity: '',
      ShiftDate: new Date().toISOString(),
      Donations10: data.d10,
      Donations20: data.d20,
      Donations25: data.d25,
      Donations30: data.d30,
      Donations35: data.d35,
      Donations40: data.d40,
      Notes: data.notes,
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
    `Welcome ${ctx.from.first_name || 'there'}!\n\nYour Telegram ID is:\n${ctx.from.id}\n\nUse /submit to submit your end-of-shift report.`
  );
});

bot.command('submit', (ctx) => {
  sessions[ctx.from.id] = { step: 'd10', data: {} };
  ctx.reply('How many $10 donations did you get?');
});

bot.on('text', async (ctx) => {
  const userId = ctx.from.id;
  const session = sessions[userId];

  if (!session) return;

  const text = ctx.message.text.trim();

  if (text.toLowerCase() === 'cancel') {
    delete sessions[userId];
    return ctx.reply('Submission cancelled.');
  }

  const askNumber = (value) => {
    const num = Number(value);
    return Number.isInteger(num) && num >= 0 ? num : null;
  };

  try {
    switch (session.step) {
      case 'd10':
        session.data.d10 = askNumber(text);
        if (session.data.d10 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'd20';
        return ctx.reply('How many $20 donations did you get?');

      case 'd20':
        session.data.d20 = askNumber(text);
        if (session.data.d20 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'd25';
        return ctx.reply('How many $25 donations did you get?');

      case 'd25':
        session.data.d25 = askNumber(text);
        if (session.data.d25 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'd30';
        return ctx.reply('How many $30 donations did you get?');

      case 'd30':
        session.data.d30 = askNumber(text);
        if (session.data.d30 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'd35';
        return ctx.reply('How many $35 donations did you get?');

      case 'd35':
        session.data.d35 = askNumber(text);
        if (session.data.d35 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'd40';
        return ctx.reply('How many $40 donations did you get?');

      case 'd40':
        session.data.d40 = askNumber(text);
        if (session.data.d40 === null) return ctx.reply('Please enter a valid number.');
        session.step = 'notes';
        return ctx.reply('Any notes? Type "none" if no notes.');

      case 'notes':
        session.data.notes = text;
        session.step = 'confirm';
        return ctx.reply(
          `Confirm submission:\n\n$10: ${session.data.d10}\n$20: ${session.data.d20}\n$25: ${session.data.d25}\n$30: ${session.data.d30}\n$35: ${session.data.d35}\n$40: ${session.data.d40}\nNotes: ${session.data.notes}\n\nType YES to submit or CANCEL to cancel.`
        );

      case 'confirm':
        if (text.toLowerCase() !== 'yes') {
          return ctx.reply('Type YES to submit or CANCEL to cancel.');
        }

        await createSubmission(ctx, session.data);
        delete sessions[userId];
        return ctx.reply('✅ End-of-shift submission saved successfully.');
    }
  } catch (error) {
    console.log(error.response?.data || error.message);
    return ctx.reply('❌ Submission failed. Please contact management.');
  }
});

bot.command('help', (ctx) => {
  ctx.reply('Commands:\n/start\n/submit\n/help');
});

bot.command('columns', async (ctx) => {
  try {
    const token = await getGraphToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/columns`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const columns = response.data.value
      .map(col => `${col.displayName} = ${col.name}`)
      .join('\n');

    ctx.reply(columns);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('Could not get columns.');
  }
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
