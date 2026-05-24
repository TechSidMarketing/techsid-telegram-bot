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

bot.command('findlists', async (ctx) => {
  try {
    const token = await getGraphToken();

    const siteResponse = await axios.get(
      'https://graph.microsoft.com/v1.0/sites/techsidmktg-my.sharepoint.com:/personal/khawar_siddiqui_techsidmktg_com',
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const siteId = siteResponse.data.id;

    const listsResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const lists = listsResponse.data.value
      .map((list) => `${list.displayName} = ${list.id}`)
      .join('\n\n');

    ctx.reply(`✅ Site ID:\n${siteId}\n\n✅ Lists Found:\n${lists}`);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Could not find SharePoint lists. Check permissions or site URL.');
  }
});

bot.command('help', (ctx) => {
  ctx.reply('Commands:\n/start\n/microsoft\n/findlists');
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
