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

function cleanText(value) {

  if (Array.isArray(value)) {
    return value.join(', ');
  }

  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function normalize(value) {

  return cleanText(value)
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

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

  return response.data.value.find(
    user =>
      user.fields.TelegramUserID === String(telegramId) &&
      user.fields.Active === true
  ) || null;
}

async function getAnyBotUser(telegramId) {

  const token = await getGraphToken();

  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items?expand=fields`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  return response.data.value.find(
    user =>
      user.fields.TelegramUserID === String(telegramId)
  ) || null;
}

async function createPendingBotUser(ctx) {

  const token = await getGraphToken();

  const fullName =
    `${ctx.from.first_name || ''} ${ctx.from.last_name || ''}`.trim();

  const payload = {
    fields: {
      Title: fullName || ctx.from.username || 'Pending User',
      TelegramUserID: String(ctx.from.id)
    }
  };

  await axios.post(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items`,
    payload,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
}

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

  return response.data.value.some(
    item =>
      item.fields.TelegramUserID === String(telegramId) &&
      item.fields.ShiftDate >= startOfDay &&
      item.fields.ShiftDate < endOfDay
  );
}

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
      Title: `Shift submission - ${cleanText(userData.RepName)}`,
      RepName: cleanText(userData.RepName),
      RepEmail: cleanText(userData.RepEmail),
      TelegramUserID: String(ctx.from.id),
      ShiftDate: new Date().toISOString(),
      TLName: cleanText(userData.TLName),
      Market_x002f_City: cleanText(userData.MarketCity),

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

bot.command('siteid', async (ctx) => {

  try {

    const token = await getGraphToken();

    const response = await axios.get(
      'https://graph.microsoft.com/v1.0/sites/techsidmktg.sharepoint.com:/sites/Globalfaces',
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    ctx.reply(`SITE ID:\n\n${response.data.id}`);

  } catch (error) {

    console.log(error.response?.data || error.message);

    ctx.reply('❌ Could not get site ID.');
  }
});

bot.start(async (ctx) => {

  const activeUser = await getBotUser(ctx.from.id);

  if (activeUser) {

    return ctx.reply(
`✅ Welcome ${activeUser.fields.LinkTitle}!

Use /submit to submit your shift report.

Use /help to see all commands.`
    );
  }

  const existingUser = await getAnyBotUser(ctx.from.id);

  if (existingUser) {

    return ctx.reply(
`⏳ Your access request is pending approval.

Management must activate your account before you can use the bot.`
    );
  }

  try {

    await createPendingBotUser(ctx);

    return ctx.reply(
`✅ Access request submitted successfully.

Your Telegram ID:
${ctx.from.id}

Management will approve your account shortly.`
    );

  } catch (error) {

    console.log(error.response?.data || error.message);

    return ctx.reply(
`❌ Could not submit access request.

Please contact management and provide this Telegram ID:
${ctx.from.id}`
    );
  }
});

bot.command('help', (ctx) => {

  ctx.reply(
`📋 Commands

/start - Start the bot or request access
/submit - Submit end-of-shift report
/mysales - View your sales today
/mytablet - View assigned tablet
/teamtoday - TL only: View today's team sales
/leaderboard - View today's leaderboard
/findlists - Show Microsoft Lists
/tabletcolumns - Show tablet list columns
/siteid - Show current SharePoint Site ID
/help - View help menu`
  );
});

bot.command('findlists', async (ctx) => {

  try {

    const token = await getGraphToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const lists = response.data.value
      .map(list => `${list.displayName} = ${list.id}`)
      .join('\n\n');

    ctx.reply(lists);

  } catch (error) {

    console.log(error.response?.data || error.message);

    ctx.reply('❌ Could not find lists.');
  }
});

bot.command('tabletcolumns', async (ctx) => {

  try {

    const token = await getGraphToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.TABLET_INVENTORY_LIST_ID}/columns`,
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

    ctx.reply('❌ Could not get tablet columns.');
  }
});

bot.command('mytablet', async (ctx) => {

  try {

    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

    const repName = normalize(user.fields.LinkTitle);

    const token = await getGraphToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.TABLET_INVENTORY_LIST_ID}/items?expand=fields`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const tablet = response.data.value.find(item => {

      const holder = normalize(item.fields.CurrentHolder);

      return (
        holder === repName ||
        holder.includes(repName) ||
        repName.includes(holder)
      );
    });

    if (!tablet) {
      return ctx.reply('📱 No tablet assigned to you.');
    }

    const fields = tablet.fields;

    const message =
`📱 My Assigned Tablet

Tablet ID: ${cleanText(fields.LinkTitle)}

Serial Number: ${cleanText(fields.SerialNumber)}

Status: ${cleanText(fields.Status)}

Condition: ${cleanText(fields.Condition)}

Accessories: ${cleanText(fields.Accessories)}

Manager: ${cleanText(fields.Manager)}

Market: ${cleanText(fields.Market)}

Lead Accepted: ${fields.LeadAccepted ? 'Yes' : 'No'}

Transfer Date: ${cleanText(fields.TransferDate)}

Power On: ${fields.PowerOn ? 'Yes' : 'No'}

Last Action By: ${cleanText(fields.LastActionBy)}`;

    ctx.reply(message);

  } catch (error) {

    console.log(error.response?.data || error.message);

    ctx.reply('❌ Failed to load tablet details.');
  }
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));

process.once('SIGTERM', () => bot.stop('SIGTERM'));
