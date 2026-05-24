require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);

bot.start((ctx) => {
  const firstName = ctx.from.first_name || 'there';
  const userId = ctx.from.id;

  ctx.reply(
    `Welcome ${firstName}!\n\nYour Telegram ID is:\n${userId}\n\nPlease send this ID to management for activation.`
  );
});

bot.command('help', (ctx) => {
  ctx.reply('TechSid Operations Bot is active.');
});

bot.command('microsoft', async (ctx) => {
  try {
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

    if (tokenResponse.data.access_token) {
      ctx.reply('✅ Microsoft connection successful.');
    } else {
      ctx.reply('❌ Microsoft connection failed.');
    }
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Microsoft connection failed. Check Railway logs.');
  }
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
