require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);

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

async function createTestSubmission(ctx) {
  const token = await getGraphToken();

  const payload = {
    fields: {
      Title: `Test submission from ${ctx.from.first_name || 'Telegram User'}`
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
  const firstName = ctx.from.first_name || 'there';
  const userId = ctx.from.id;

  ctx.reply(
    `Welcome ${firstName}!\n\nYour Telegram ID is:\n${userId}\n\nPlease send this ID to management for activation.`
  );
});

bot.command('microsoft', async (ctx) => {
  try {
    await getGraphToken();
    ctx.reply('✅ Microsoft connection successful.');
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Microsoft connection failed.');
  }
});

bot.command('testsubmission', async (ctx) => {
  try {
    await createTestSubmission(ctx);
    ctx.reply('✅ Test submission created in Microsoft List.');
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Test submission failed. Check Railway logs.');
  }
});

bot.command('help', (ctx) => {
  ctx.reply('Commands:\n/start\n/microsoft\n/testsubmission');
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
