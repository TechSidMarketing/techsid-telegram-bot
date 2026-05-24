require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);
const sessions = {};

async function getGraphToken() {
  const res = await axios.post(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    }),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );

  return res.data.access_token;
}

function cleanText(value) {
  if (Array.isArray(value)) return value.join(', ');
  if (value === null || value === undefined) return '';
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

  const res = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items?expand=fields`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  return res.data.value.find(
    user =>
      cleanText(user.fields.TelegramUserID) === String(telegramId) &&
      user.fields.Active === true
  ) || null;
}

async function getAnyBotUser(telegramId) {
  const token = await getGraphToken();

  const res = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items?expand=fields`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  return res.data.value.find(
    user => cleanText(user.fields.TelegramUserID) === String(telegramId)
  ) || null;
}

async function createPendingBotUser(ctx) {
  const token = await getGraphToken();

  const fullName =
    `${ctx.from.first_name || ''} ${ctx.from.last_name || ''}`.trim();

  await axios.post(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.BOT_USERS_LIST_ID}/items`,
    {
      fields: {
        Title: fullName || ctx.from.username || 'Pending User',
        TelegramUserID: String(ctx.from.id)
      }
    },
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

  const res = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  return res.data.value.some(
    item =>
      cleanText(item.fields.TelegramUserID) === String(telegramId) &&
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

  await axios.post(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items`,
    {
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
    },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
}

bot.start(async (ctx) => {
  const activeUser = await getBotUser(ctx.from.id);

  if (activeUser) {
    return ctx.reply(
`✅ Welcome ${activeUser.fields.Title || activeUser.fields.LinkTitle}!

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
/columns - Show sales list columns
/tabletcolumns - Show tablet list columns
/siteid - Show SharePoint site ID
/help - View help menu`
  );
});

bot.command('siteid', async (ctx) => {
  try {
    const token = await getGraphToken();

    const res = await axios.get(
      'https://graph.microsoft.com/v1.0/sites/techsidmktg.sharepoint.com:/sites/Globalfaces',
      { headers: { Authorization: `Bearer ${token}` } }
    );

    ctx.reply(`SITE ID:\n\n${res.data.id}`);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Could not get site ID.');
  }
});

bot.command('findlists', async (ctx) => {
  try {
    const token = await getGraphToken();

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const lists = res.data.value
      .map(list => `${list.displayName} = ${list.id}`)
      .join('\n\n');

    ctx.reply(lists);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Could not find lists.');
  }
});

bot.command('columns', async (ctx) => {
  try {
    const token = await getGraphToken();

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/columns`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const columns = res.data.value
      .map(col => `${col.displayName} = ${col.name}`)
      .join('\n');

    ctx.reply(columns);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Could not get columns.');
  }
});

bot.command('tabletcolumns', async (ctx) => {
  try {
    const token = await getGraphToken();

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.TABLET_INVENTORY_LIST_ID}/columns`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const columns = res.data.value
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

    const repName = normalize(
      user.fields.Title ||
      user.fields.LinkTitle
    );

    const token = await getGraphToken();

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.TABLET_INVENTORY_LIST_ID}/items?expand=fields`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const tablet = res.data.value.find(item => {
      const holder = normalize(
        item.fields.CurrentHolder ||
        item.fields.CurrentRepHolder
      );

      return (
        holder === repName ||
        holder.includes(repName) ||
        repName.includes(holder)
      );
    });

    if (!tablet) {
      return ctx.reply('📱 No tablet assigned to you.');
    }

    const f = tablet.fields;

    ctx.reply(
`📱 My Assigned Tablet

Tablet ID: ${cleanText(f.LinkTitle)}
Serial Number: ${cleanText(f.SerialNumber)}
Status: ${cleanText(f.Status)}
Condition: ${cleanText(f.Condition)}
Accessories: ${cleanText(f.Accessories)}
Manager: ${cleanText(f.Manager)}
Market: ${cleanText(f.Market)}
Lead Accepted: ${f.LeadAccepted ? 'Yes' : 'No'}
Transfer Date: ${cleanText(f.TransferDate)}
Power On: ${f.PowerOn ? 'Yes' : 'No'}
Last Action By: ${cleanText(f.LastActionBy)}`
    );
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Failed to load tablet details.');
  }
});

bot.command('mysales', async (ctx) => {
  try {
    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

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

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const mySubmissions = res.data.value.filter(
      item =>
        cleanText(item.fields.TelegramUserID) === String(ctx.from.id) &&
        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
    );

    if (mySubmissions.length === 0) {
      return ctx.reply('You have not submitted any sales today yet.');
    }

    const item = mySubmissions[0].fields;

    const actualTotal =
      (item._x0024_10Donations || 0) +
      (item._x0024_20Donations || 0) +
      (item._x0024_25Donations || 0) +
      (item._x0024_30Donations || 0) +
      (item._x0024_35Donations || 0) +
      (item._x0024_40Donations || 0);

    ctx.reply(
`📊 My Sales Today

$10: ${item._x0024_10Donations || 0}
$20: ${item._x0024_20Donations || 0}
$25: ${item._x0024_25Donations || 0}
$30: ${item._x0024_30Donations || 0}
$35: ${item._x0024_35Donations || 0}
$40: ${item._x0024_40Donations || 0}

🔥 Total Donations: ${item.TotalDonations || actualTotal}

📌 Status: ${item.Status || 'Submitted'}`
    );
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Failed to load your sales.');
  }
});

bot.command('leaderboard', async (ctx) => {
  try {
    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

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

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const ranked = res.data.value
      .filter(item =>
        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
      )
      .map(item => ({
        rep: cleanText(item.fields.RepName) || 'Unknown Rep',
        total: item.fields.TotalDonations || 0
      }))
      .sort((a, b) => b.total - a.total);

    if (ranked.length === 0) {
      return ctx.reply('No submissions today yet.');
    }

    let message = '🏆 Today’s Leaderboard\n\n';

    ranked.forEach((item, index) => {
      message += `${index + 1}. ${item.rep} - ${item.total}\n`;
    });

    ctx.reply(message);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Failed to load leaderboard.');
  }
});

bot.command('teamtoday', async (ctx) => {
  try {
    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

    const roles = user.fields.Role || [];

    if (!roles.includes('TL')) {
      return ctx.reply('❌ Only Team Leads can use this command.');
    }

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

    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.REP_SUBMISSIONS_LIST_ID}/items?expand=fields`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const tlName = cleanText(user.fields.Title || user.fields.LinkTitle);

    const todaySubs = res.data.value.filter(
      item =>
        cleanText(item.fields.TLName) === tlName &&
        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
    );

    if (todaySubs.length === 0) {
      return ctx.reply('No team submissions today.');
    }

    let total = 0;

    let message = '📊 Team Today\n\n';

    todaySubs.forEach(item => {
      const repTotal = item.fields.TotalDonations || 0;

      total += repTotal;

      message += `${cleanText(item.fields.RepName)} - ${repTotal}\n`;
    });

    message += `\n🔥 Team Total: ${total}`;

    ctx.reply(message);
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Failed to load team report.');
  }
});

bot.command('submit', async (ctx) => {
  try {
    const user = await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ You are not authorized to submit.');
    }

    const alreadySubmitted = await hasSubmittedToday(ctx.from.id);

    if (alreadySubmitted) {
      return ctx.reply('❌ You have already submitted your shift report today.');
    }

    sessions[ctx.from.id] = {
      step: 'd10',
      data: {},
      user: {
        RepName: user.fields.Title || user.fields.LinkTitle || '',
        RepEmail: user.fields.Email || '',
        TLName: user.fields.TL_x002f_MangerName || '',
        MarketCity: user.fields.Market_x002f_City || ''
      }
    };

    ctx.reply('How many $10 donations did you get?');
  } catch (error) {
    console.log(error.response?.data || error.message);
    ctx.reply('❌ Could not start submission.');
  }
});

bot.on('text', async (ctx) => {
  const userId = ctx.from.id;

  const text = ctx.message.text.trim();

  if (text.startsWith('/')) {
    return;
  }

  const session = sessions[userId];

  if (!session) {
    return;
  }

  if (text.toLowerCase() === 'cancel') {
    delete sessions[userId];
    return ctx.reply('❌ Submission cancelled.');
  }

  const askNumber = value => {
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
        session.step = 'confirm';

        return ctx.reply(
`Confirm submission:

$10: ${session.data.d10}
$20: ${session.data.d20}
$25: ${session.data.d25}
$30: ${session.data.d30}
$35: ${session.data.d35}
$40: ${session.data.d40}

🔥 Total Donations: ${
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
        if (text.toLowerCase() !== 'yes') {
          return ctx.reply('Type YES to submit or CANCEL to cancel.');
        }

        await createSubmission(ctx, session.data, session.user);

        delete sessions[userId];

        return ctx.reply('✅ End-of-shift submission saved successfully.');
    }
  } catch (error) {
    console.log(error.response?.data || error.message);
    return ctx.reply('❌ Submission failed. Please contact management.');
  }
});

bot.catch((err) => {
  console.log('Bot error:', err);
});

bot.launch();

console.log('TechSid Telegram Bot is running...');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
