require('dotenv').config();

const { Telegraf } = require('telegraf');
const axios = require('axios');

const bot = new Telegraf(process.env.BOT_TOKEN);

const sessions = {};

// ======================
// CONFIG
// ======================

const EMPLOYEE_ID_FIELD =
  process.env.EMPLOYEE_ID_FIELD || 'EmployeeNumber';

// ======================
// HELPERS
// ======================

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

function normalizeEmployeeId(value) {

  return cleanText(value)
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function formatAccessories(value) {

  if (Array.isArray(value)) {
    return value.join(', ');
  }

  return cleanText(value);
}

function getTodayRange() {

  const today = new Date();

  return {

    startOfDay:
      new Date(
        today.getFullYear(),
        today.getMonth(),
        today.getDate()
      ).toISOString(),

    endOfDay:
      new Date(
        today.getFullYear(),
        today.getMonth(),
        today.getDate() + 1
      ).toISOString()
  };
}

// ======================
// GRAPH TOKEN
// ======================

async function getGraphToken() {

  const response =
    await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,

      new URLSearchParams({

        client_id:
          process.env.CLIENT_ID,

        client_secret:
          process.env.CLIENT_SECRET,

        scope:
          'https://graph.microsoft.com/.default',

        grant_type:
          'client_credentials'
      }),

      {
        headers: {
          'Content-Type':
            'application/x-www-form-urlencoded'
        }
      }
    );

  return response.data.access_token;
}

// ======================
// SHAREPOINT HELPERS
// ======================

async function getListItems(listId) {

  const token =
    await getGraphToken();

  const response =
    await axios.get(

      `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items?expand=fields`,

      {
        headers: {
          Authorization:
            `Bearer ${token}`
        }
      }
    );

  return response.data.value;
}

async function updateListItemFields(
  listId,
  itemId,
  fields
) {

  const token =
    await getGraphToken();

  await axios.patch(

    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items/${itemId}/fields`,

    fields,

    {
      headers: {

        Authorization:
          `Bearer ${token}`,

        'Content-Type':
          'application/json'
      }
    }
  );
}

// ======================
// BOT USERS
// ======================

async function getAllBotUsers() {

  return await getListItems(
    process.env.BOT_USERS_LIST_ID
  );
}

async function getBotUser(telegramId) {

  const users =
    await getAllBotUsers();

  return users.find(user =>

    cleanText(
      user.fields.TelegramUserID
    ) === String(telegramId)

  ) || null;
}

async function getAnyBotUser(
  telegramId
) {

  const users =
    await getAllBotUsers();

  return users.find(user =>

    cleanText(
      user.fields.TelegramUserID
    ) === String(telegramId)

  ) || null;
}

async function findBotUserByRegistration(
  firstName,
  lastName,
  employeeId
) {

  const users =
    await getAllBotUsers();

  const submittedFullName =
    normalize(
      `${firstName} ${lastName}`
    );

  const submittedEmployeeId =
    normalizeEmployeeId(
      employeeId
    );

  return users.find(user => {

    const fullName =
      normalize(
        user.fields.Title ||
        user.fields.LinkTitle
      );

    const listEmployeeId =
      normalizeEmployeeId(
        user.fields[EMPLOYEE_ID_FIELD]
      );

    return (
      fullName === submittedFullName &&
      listEmployeeId === submittedEmployeeId
    );

  }) || null;
}

async function registerBotUser(
  botUserItem,
  telegramId
) {

  await updateListItemFields(

    process.env.BOT_USERS_LIST_ID,

    botUserItem.id,

    {
      TelegramUserID:
        String(telegramId),

      Active: true
    }
  );
}

// ======================
// START COMMAND
// ======================

bot.start(async (ctx) => {

  try {

    const activeUser =
      await getBotUser(ctx.from.id);

    if (activeUser) {

      return ctx.reply(
`✅ Welcome ${activeUser.fields.Title || activeUser.fields.LinkTitle}!

Use /submit to submit your shift report.
Use /help to see all commands.`
      );
    }

    const existingUser =
      await getAnyBotUser(ctx.from.id);

    if (existingUser) {

      return ctx.reply(
`⏳ Your Telegram account is already linked but not active.

Please contact your administrator.`
      );
    }

    sessions[ctx.from.id] = {

      type: 'registerUser',

      step: 'firstName',

      data: {}
    };

    return ctx.reply(
`Welcome to TechSidBot.

To verify your access, please enter your first name.`
    );

  } catch (error) {

    console.log(
      error.response?.data || error.message
    );

    return ctx.reply(
      '❌ Could not start registration. Please contact your administrator.'
    );
  }
});

// ======================
// HELP COMMAND
// ======================

bot.command('help', (ctx) => {

  ctx.reply(
`📋 Commands

/start - Start bot
/sale - Log live sale
/submit - Submit end-of-day report
/mysales - View today's sales
/leaderboard - Live sales leaderboard
/teamtoday - Team report
/mytablet - View assigned tablets
/assigntablet - Assign tablet
/accepttablet - Accept assigned tablet
/help - Help menu`
  );
});

// ======================
// SALES FUNCTIONS
// ======================

async function hasSubmittedToday(telegramId) {

  const {
    startOfDay,
    endOfDay
  } = getTodayRange();

  const submissions =
    await getListItems(
      process.env.REP_SUBMISSIONS_LIST_ID
    );

  return submissions.some(item =>

    cleanText(
      item.fields.TelegramUserID
    ) === String(telegramId) &&

    item.fields.ShiftDate >= startOfDay &&
    item.fields.ShiftDate < endOfDay
  );
}

async function createSubmission(
  ctx,
  data,
  userData
) {

  const token =
    await getGraphToken();

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

        Title:
          `Shift submission - ${cleanText(userData.RepName)}`,

        RepName:
          cleanText(userData.RepName),

        RepEmail:
          cleanText(userData.RepEmail),

        TelegramUserID:
          String(ctx.from.id),

        TLName:
          cleanText(userData.TLName),

        Market_x002f_City:
          cleanText(userData.MarketCity),

        ShiftDate:
          new Date().toISOString(),

        _x0024_10Donations:
          data.d10,

        _x0024_20Donations:
          data.d20,

        _x0024_25Donations:
          data.d25,

        _x0024_30Donations:
          data.d30,

        _x0024_35Donations:
          data.d35,

        _x0024_40Donations:
          data.d40,

        TotalDonations:
          totalDonations,

        Status:
          'Submitted'
      }
    },

    {
      headers: {
        Authorization:
          `Bearer ${token}`,

        'Content-Type':
          'application/json'
      }
    }
  );
}

// ======================
// SUBMIT SALES
// ======================

bot.command('submit', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

    const alreadySubmitted =
      await hasSubmittedToday(ctx.from.id);

    if (alreadySubmitted) {
      return ctx.reply(
        '❌ You have already submitted your shift report today.'
      );
    }

    const {
      startOfDay,
      endOfDay
    } = getTodayRange();

    const liveSales =
      await getListItems(
        process.env.LIVE_SALES_LIST_ID
      );

    const myLiveSales =
      liveSales.filter(item =>
        cleanText(item.fields.TelegramUserID) === String(ctx.from.id) &&
        item.fields.Sale >= startOfDay &&
        item.fields.Sale < endOfDay
      );

    const totals = {
      d10: 0,
      d20: 0,
      d25: 0,
      d30: 0,
      d35: 0,
      d40: 0
    };

    myLiveSales.forEach(item => {

      const amount =
        cleanText(item.fields.DonationAmount)
          .replace('$', '')
          .trim();

      if (amount === '10') totals.d10 += 1;
      if (amount === '20') totals.d20 += 1;
      if (amount === '25') totals.d25 += 1;
      if (amount === '30') totals.d30 += 1;
      if (amount === '35') totals.d35 += 1;
      if (amount === '40') totals.d40 += 1;
    });

    sessions[ctx.from.id] = {

      type:
        'confirmEOD',

      step:
        'confirm',

      data:
        totals,

      user: {

        RepName:
          user.fields.Title ||
          user.fields.LinkTitle ||
          '',

        RepEmail:
          user.fields.Email ||
          '',

        TLName:
          user.fields.TL_x002f_MangerName ||
          '',

        MarketCity:
          user.fields.Market_x002f_City ||
          ''
      }
    };

    const totalDonations =
      totals.d10 +
      totals.d20 +
      totals.d25 +
      totals.d30 +
      totals.d35 +
      totals.d40;

    return ctx.reply(
`📋 End-of-Day Summary

$10: ${totals.d10}
$20: ${totals.d20}
$25: ${totals.d25}
$30: ${totals.d30}
$35: ${totals.d35}
$40: ${totals.d40}

🔥 Total Donations:
${totalDonations}

Type YES to submit
or EDIT to adjust.`
    );

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Could not start end-of-day submission.'
    );
  }
});

// ======================
// MY SALES
// ======================

bot.command('mysales', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    const {
      startOfDay,
      endOfDay
    } = getTodayRange();

    const submissions =
      await getListItems(
        process.env.REP_SUBMISSIONS_LIST_ID
      );

    const mySubmissions =
      submissions.filter(item =>

        cleanText(
          item.fields.TelegramUserID
        ) === String(ctx.from.id) &&

        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
      );

    if (mySubmissions.length === 0) {

      return ctx.reply(
        'You have not submitted any sales today yet.'
      );
    }

    const item =
      mySubmissions[0].fields;

    return ctx.reply(
`📊 My Sales Today

$10: ${item._x0024_10Donations || 0}
$20: ${item._x0024_20Donations || 0}
$25: ${item._x0024_25Donations || 0}
$30: ${item._x0024_30Donations || 0}
$35: ${item._x0024_35Donations || 0}
$40: ${item._x0024_40Donations || 0}

🔥 Total Donations:
${item.TotalDonations || 0}

📌 Status:
${item.Status || 'Submitted'}`
    );

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Failed to load your sales.'
    );
  }
});

// ======================
// CREATE LIST ITEM
// ======================

async function createListItem(listId, fields) {

  const token =
    await getGraphToken();

  await axios.post(
    `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`,
    {
      fields
    },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
}

// ======================
// CREATE LIVE SALE
// ======================

async function createLiveSale(
  ctx,
  donationAmount,
  user
) {

  const fields = {

  Title:
    cleanText(
      user.fields.LinkTitle ||
      user.fields.Title
    ),

  Sale:
    new Date().toISOString(),

  RepEmail:
    cleanText(
      user.fields.Email
    ),

  TelegramUserID:
    String(ctx.from.id),

  TL_x002f_ManagerName:
    cleanText(
      user.fields.TL_x002f_MangerName
    ),

  Market_x002f_City:
    cleanText(
      user.fields.Market_x002f_City
    ),

  DonationAmount:
  `$${donationAmount}`
};

  await createListItem(
    process.env.LIVE_SALES_LIST_ID,
    fields
  );
}

// ======================
// SALE COMMAND
// ======================

bot.command('sale', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {

      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    sessions[ctx.from.id] = {

      type:
        'liveSale',

      step:
        'amount',

      data:
        {},

      user
    };

    return ctx.reply(
`Enter donation amount:

10
20
25
30
35
40`
    );

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Could not start sale entry.'
    );
  }
});

// ======================
// TEAM TODAY
// ======================

bot.command('teamtoday', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {

      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    const role =
  normalize(user.fields.Role);

if (
  !role.includes('tl') &&
  !role.includes('admin')
) {

      return ctx.reply(
        '❌ Only TL/Admin can use this command.'
      );
    }

    const {
      startOfDay,
      endOfDay
    } = getTodayRange();

    const tlName =
      cleanText(
        user.fields.Title ||
        user.fields.LinkTitle
      );

    const submissions =
      await getListItems(
        process.env.REP_SUBMISSIONS_LIST_ID
      );

    const teamSubs =
      submissions.filter(item =>

        cleanText(
          item.fields.TLName
        ) === tlName &&

        item.fields.ShiftDate >= startOfDay &&
        item.fields.ShiftDate < endOfDay
      );

    if (teamSubs.length === 0) {

      return ctx.reply(
        'No team submissions today.'
      );
    }

    let total = 0;

    let message =
      '📊 Team Today\n\n';

    teamSubs.forEach(item => {

      const repTotal =
        item.fields.TotalDonations || 0;

      total += repTotal;

      message +=
`${cleanText(item.fields.RepName)} - ${repTotal}\n`;
    });

    message +=
`\n🔥 Team Total: ${total}`;

    return ctx.reply(message);

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Failed to load team report.'
    );
  }
});

// ======================
// TABLET FUNCTIONS
// ======================

async function assignTablet(
  tabletId,
  assigneeUser,
  assignType,
  assignedByName
) {

  const tablets =
    await getListItems(
      process.env.TABLET_INVENTORY_LIST_ID
    );

  const tablet =
    tablets.find(item =>

      normalize(
        item.fields.LinkTitle
      ) === normalize(tabletId)
    );

  if (!tablet) {

    throw new Error(
      'Tablet not found'
    );
  }

  const assigneeName =
    cleanText(
      assigneeUser.fields.Title ||
      assigneeUser.fields.LinkTitle
    );

  const assigneeMarket =
    cleanText(
      assigneeUser.fields.Market_x002f_City
    );

  const assigneeManager =
    cleanText(
      assigneeUser.fields.TL_x002f_MangerName
    );

  const updateFields = {

    Status:
      'Pending Acceptance',

    Market:
      assigneeMarket,

    LeadAccepted:
      false,

    TransferDate:
      new Date().toISOString(),

    LastActionBy:
      assignedByName
  };

  if (assignType === 'TL') {

    updateFields.CurrentHolder = '';
    updateFields.Manager = assigneeName;

  } else {

    updateFields.CurrentHolder =
      assigneeName;

    updateFields.Manager =
      assigneeManager || assignedByName;
  }

  await updateListItemFields(

    process.env.TABLET_INVENTORY_LIST_ID,

    tablet.id,

    updateFields
  );
}

async function acceptTablet(
  tabletItemId,
  tabletData,
  repName
) {

  const accessories =
`Charging Cable: ${tabletData.chargingCable}
Charging Block: ${tabletData.chargingBlock}`;

  await updateListItemFields(

    process.env.TABLET_INVENTORY_LIST_ID,

    tabletItemId,

    {

      Status:
        'Active',

      LeadAccepted:
        true,

      Condition:
        tabletData.condition,

      Accessories:
        accessories,

      PowerOn:
        tabletData.powerOn,

      Notes:
        tabletData.notes,

      LastActionBy:
        repName,

      TransferDate:
        new Date().toISOString()
    }
  );
}

// ======================
// MY TABLET
// ======================

bot.command('mytablet', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {

      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    const repName =
      normalize(
        user.fields.Title ||
        user.fields.LinkTitle
      );

    const tablets =
      await getListItems(
        process.env.TABLET_INVENTORY_LIST_ID
      );

    const assignedTablets =
      tablets.filter(item => {

        const holder =
          normalize(
            item.fields.CurrentHolder
          );

        return (
          holder === repName ||
          holder.includes(repName) ||
          repName.includes(holder)
        );

      });

    if (assignedTablets.length === 0) {

      return ctx.reply(
        '📱 No tablet assigned to you.'
      );
    }

    let message =
`📱 My Assigned Tablets (${assignedTablets.length})

`;

    assignedTablets.forEach((tablet, index) => {

      const f =
        tablet.fields;

      message +=
`${index + 1}. Tablet ID: ${cleanText(f.LinkTitle)}

Serial Number: ${cleanText(f.SerialNumber)}

Status: ${cleanText(f.Status)}

Condition: ${cleanText(f.Condition)}

Accessories: ${formatAccessories(f.Accessories)}

Manager: ${cleanText(f.Manager)}

Market: ${cleanText(f.Market)}

Lead Accepted:
${f.LeadAccepted ? 'Yes' : 'No'}

Transfer Date:
${cleanText(f.TransferDate)}

Power On:
${f.PowerOn ? 'Yes' : 'No'}

Notes:
${cleanText(f.Notes)}

-------------------------

`;
    });

    return ctx.reply(message);

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Failed to load tablet details.'
    );
  }
});

// ======================
// ASSIGN TABLET
// ======================

bot.command('assigntablet', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {

      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    const role =
  normalize(user.fields.Role);

if (
  !role.includes('tl') &&
  !role.includes('admin')
) {

      return ctx.reply(
        '❌ Only TL/Admin can assign tablets.'
      );
    }

    sessions[ctx.from.id] = {

      type:
        'assignTablet',

      step:
        'tabletId',

      data:
        {},

      managerName:
        cleanText(
          user.fields.Title ||
          user.fields.LinkTitle
        )
    };

    return ctx.reply(
      'Enter the Tablet ID you want to assign.'
    );

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Could not start tablet assignment.'
    );
  }
});

// ======================
// ACCEPT TABLET
// ======================

bot.command('accepttablet', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {

      return ctx.reply(
        '❌ Unauthorized.'
      );
    }

    const repName =
      normalize(
        user.fields.Title ||
        user.fields.LinkTitle
      );

    const tablets =
      await getListItems(
        process.env.TABLET_INVENTORY_LIST_ID
      );

    const pendingTablets =
      tablets.filter(item => {

        const holder =
          normalize(
            item.fields.CurrentHolder
          );

        return (
          holder === repName &&
          item.fields.Status === 'Pending Acceptance'
        );

      });

    if (pendingTablets.length === 0) {

      return ctx.reply(
        '❌ No pending tablets found.'
      );
    }

    // ======================
    // ONLY 1 TABLET
    // ======================

    if (pendingTablets.length === 1) {

      const tablet =
        pendingTablets[0];

      sessions[ctx.from.id] = {

        type:
          'acceptTablet',

        step:
          'condition',

        tabletItemId:
          tablet.id,

        tabletId:
          cleanText(
            tablet.fields.LinkTitle
          ),

        data:
          {}
      };

      return ctx.reply(
`📱 Accept Tablet

Tablet ID:
${cleanText(tablet.fields.LinkTitle)}

Current Condition:
${cleanText(tablet.fields.Condition)}

Accessories:
${formatAccessories(tablet.fields.Accessories)}

Is the tablet condition acceptable?

Reply:
YES
or
NO`
      );
    }

    // ======================
    // MULTIPLE TABLETS
    // ======================

    sessions[ctx.from.id] = {

      type:
        'acceptTablet',

      step:
        'selectTablet',

      pendingTablets,

      data:
        {}
    };

    let message =
`📱 Pending Tablets (${pendingTablets.length})

`;

    pendingTablets.forEach((tablet, index) => {

      const f =
        tablet.fields;

      message +=
`${index + 1}. Tablet ID: ${cleanText(f.LinkTitle)}

Condition:
${cleanText(f.Condition)}

Accessories:
${formatAccessories(f.Accessories)}

Power On:
${f.PowerOn ? 'Yes' : 'No'}

`;
    });

    message +=
'Reply with the Tablet ID you want to accept.';

    return ctx.reply(message);

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Failed to start tablet acceptance.'
    );
  }
});

bot.command('leaderboard', async (ctx) => {

  try {

    const user =
      await getBotUser(ctx.from.id);

    if (!user) {
      return ctx.reply('❌ Unauthorized.');
    }

    const {
      startOfDay,
      endOfDay
    } = getTodayRange();

    const sales =
      await getListItems(
        process.env.LIVE_SALES_LIST_ID
      );

    const todaySales =
      sales.filter(item =>
        item.fields.Sale >= startOfDay &&
        item.fields.Sale < endOfDay
      );

    if (todaySales.length === 0) {
      return ctx.reply('No live sales logged today yet.');
    }

    const leaderboard = {};

    todaySales.forEach(item => {

      const rep =
        cleanText(item.fields.Title || item.fields.LinkTitle);

      if (!leaderboard[rep]) {
        leaderboard[rep] = 0;
      }

      leaderboard[rep] += 1;
    });

    const ranked =
      Object.entries(leaderboard)
        .map(([rep, total]) => ({ rep, total }))
        .sort((a, b) => b.total - a.total);

    let message =
      '🏆 Live Sales Leaderboard\n\n';

    ranked.forEach((item, index) => {
      message += `${index + 1}. ${item.rep} - ${item.total}\n`;
    });

    const totalSales =
      ranked.reduce((sum, item) => sum + item.total, 0);

    message += `\n🔥 Total Sales: ${totalSales}`;

    return ctx.reply(message);

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Failed to load live sales leaderboard.'
    );
  }
});

// ======================
// TEXT SESSION HANDLER
// ======================

bot.on('text', async (ctx) => {

  const userId =
    ctx.from.id;

  const text =
    ctx.message.text.trim();

  if (text.startsWith('/')) {
    return;
  }

  const session =
    sessions[userId];

  if (!session) {
    return;
  }

  if (text.toLowerCase() === 'cancel') {

    delete sessions[userId];

    return ctx.reply(
      '❌ Cancelled.'
    );
  }

  try {

    // ======================
// LIVE SALE FLOW
// ======================

if (session.type === 'liveSale') {

  if (session.step === 'amount') {

    const amount =
      Number(text);

    const allowedAmounts =
      [10, 20, 25, 30, 35, 40];

    if (!allowedAmounts.includes(amount)) {

      return ctx.reply(
        '❌ Please enter a valid amount: 10, 20, 25, 30, 35, or 40.'
      );
    }

    session.data.amount =
      amount;

    session.step =
      'confirm';

    return ctx.reply(
`Confirm sale:

Donation Amount:
$${amount}

Type YES to save
or CANCEL to cancel.`
    );
  }

  if (session.step === 'confirm') {

    if (text.toLowerCase() !== 'yes') {

      return ctx.reply(
        'Type YES to save or CANCEL to cancel.'
      );
    }

    await createLiveSale(
      ctx,
      session.data.amount,
      session.user
    );

    delete sessions[userId];

    return ctx.reply(
      '✅ Sale logged successfully.'
    );
  }
}

    // ======================
    // REGISTRATION FLOW
    // ======================

    if (session.type === 'registerUser') {

      if (session.step === 'firstName') {

        session.data.firstName =
          text;

        session.step =
          'lastName';

        return ctx.reply(
          'Please enter your last name.'
        );
      }

      if (session.step === 'lastName') {

        session.data.lastName =
          text;

        session.step =
          'employeeId';

        return ctx.reply(
          'Please enter your Employee ID.'
        );
      }

      if (session.step === 'employeeId') {

        session.data.employeeId =
          text;

        const matchedUser =
          await findBotUserByRegistration(
            session.data.firstName,
            session.data.lastName,
            session.data.employeeId
          );

        if (!matchedUser) {

          delete sessions[userId];

          return ctx.reply(
            '❌ User not found. Contact your administrator.'
          );
        }

        const existingTelegramId =
          cleanText(
            matchedUser.fields.TelegramUserID
          );

        if (
          existingTelegramId &&
          existingTelegramId !== String(ctx.from.id)
        ) {

          delete sessions[userId];

          return ctx.reply(
            '❌ This Employee ID is already linked to another Telegram account. Please contact your administrator.'
          );
        }

        await registerBotUser(
          matchedUser,
          ctx.from.id
        );

        delete sessions[userId];

        return ctx.reply(
`✅ Registration successful.

Welcome ${
  matchedUser.fields.Title ||
  matchedUser.fields.LinkTitle
}!

Use /help to see available commands.`
        );
      }
    }

    // ======================
    // ASSIGN TABLET FLOW
    // ======================

    if (session.type === 'assignTablet') {

      if (session.step === 'tabletId') {

        session.data.tabletId =
          text;

        session.step =
          'assignType';

        return ctx.reply(
`Who are you assigning this tablet to?

Type:
REP
or
TL`
        );
      }

      if (session.step === 'assignType') {

        const type =
          text.toUpperCase();

        if (
          type !== 'REP' &&
          type !== 'TL'
        ) {

          return ctx.reply(
            'Please type REP or TL.'
          );
        }

        session.data.assignType =
          type;

        session.step =
          'assigneeName';

        if (type === 'REP') {

          return ctx.reply(
            'Enter the employee ID.'
          );
        }

        return ctx.reply(
          'Enter the TL/Manager employee ID.'
        );
      }

      if (session.step === 'assigneeName') {

        const users = await getAllBotUsers();

const assigneeUser = users.find(user =>
  normalizeEmployeeId(user.fields[EMPLOYEE_ID_FIELD]) === normalizeEmployeeId(text)
);

        if (!assigneeUser) {

          return ctx.reply(
            '❌ User not found. Please contact your administrator.'
          );
        }

        session.data.assigneeUser =
          assigneeUser;

        const assigneeName =
          cleanText(
            assigneeUser.fields.Title ||
            assigneeUser.fields.LinkTitle
          );

        const assigneeMarket =
          cleanText(
            assigneeUser.fields.Market_x002f_City
          );

        const assigneeManager =
          cleanText(
            assigneeUser.fields.TL_x002f_MangerName
          );

        session.step =
          'confirm';

        return ctx.reply(
`Confirm tablet assignment:

Tablet ID:
${session.data.tabletId}

Assign To:
${assigneeName}

Assign Type:
${session.data.assignType}

Market/City:
${assigneeMarket}

Manager:
${
  session.data.assignType === 'TL'
    ? assigneeName
    : assigneeManager || session.managerName
}

Assigned By:
${session.managerName}

Type YES to assign
or CANCEL to cancel.`
        );
      }

      if (session.step === 'confirm') {

        if (
          text.toLowerCase() !== 'yes'
        ) {

          return ctx.reply(
            'Type YES to assign or CANCEL to cancel.'
          );
        }

        await assignTablet(
          session.data.tabletId,
          session.data.assigneeUser,
          session.data.assignType,
          session.managerName
        );

        delete sessions[userId];

        return ctx.reply(
          '✅ Tablet assignment saved successfully.'
        );
      }
    }

      // ======================
    // ACCEPT TABLET FLOW
    // ======================

    if (session.type === 'acceptTablet') {

      // ======================
      // SELECT TABLET
      // ======================

      if (session.step === 'selectTablet') {

        const cleanedInput =
          normalize(text)

            .replace('accept tablet id:', '')
            .replace('accept tablet:', '')
            .replace('accept tablet id', '')
            .replace('accept', '')
            .trim();

        const selectedTablet =
          session.pendingTablets.find(tablet =>

            normalize(
              tablet.fields.LinkTitle
            ) === cleanedInput
          );

        if (!selectedTablet) {

          return ctx.reply(
            '❌ Tablet not found. Please enter a valid Tablet ID from the list.'
          );
        }

        session.tabletItemId =
          selectedTablet.id;

        session.tabletId =
          cleanText(
            selectedTablet.fields.LinkTitle
          );

        session.step =
          'condition';

        return ctx.reply(
`📱 Accept Tablet

Tablet ID:
${cleanText(selectedTablet.fields.LinkTitle)}

Current Condition:
${cleanText(selectedTablet.fields.Condition)}

Accessories:
${formatAccessories(selectedTablet.fields.Accessories)}

Is the tablet condition acceptable?

Reply:
YES
or
NO`
        );
      }

      // ======================
      // CONDITION
      // ======================

      if (session.step === 'condition') {

        session.data.condition =
          text.toUpperCase() === 'YES'
            ? 'Accepted'
            : 'Issue Reported';

        session.step =
          'chargingCable';

        return ctx.reply(
`Was the charging cable provided?

Reply:
YES
or
NO`
        );
      }

      // ======================
      // CHARGING CABLE
      // ======================

      if (session.step === 'chargingCable') {

        session.data.chargingCable =
          text.toUpperCase() === 'YES'
            ? 'Yes'
            : 'No';

        session.step =
          'chargingBlock';

        return ctx.reply(
`Was the charging block provided?

Reply:
YES
or
NO`
        );
      }

      // ======================
      // CHARGING BLOCK
      // ======================

      if (session.step === 'chargingBlock') {

        session.data.chargingBlock =
          text.toUpperCase() === 'YES'
            ? 'Yes'
            : 'No';

        session.step =
          'powerOn';

        return ctx.reply(
`Does the tablet power on?

Reply:
YES
or
NO`
        );
      }

      // ======================
      // POWER ON
      // ======================

      if (session.step === 'powerOn') {

        session.data.powerOn =
          text.toUpperCase() === 'YES';

        session.step =
          'notes';

        return ctx.reply(
`Any notes?

Type your notes or type:
NONE`
        );
      }

      // ======================
      // NOTES
      // ======================

      if (session.step === 'notes') {

        session.data.notes =
          text.toUpperCase() === 'NONE'
            ? ''
            : text;

        session.step =
          'confirm';

        return ctx.reply(
`Confirm tablet acceptance?

Tablet ID:
${session.tabletId}

Condition:
${session.data.condition}

Charging Cable:
${session.data.chargingCable}

Charging Block:
${session.data.chargingBlock}

Power On:
${session.data.powerOn ? 'Yes' : 'No'}

Notes:
${session.data.notes || 'None'}

Type YES to confirm
or CANCEL to cancel.`
        );
      }

      // ======================
      // FINAL CONFIRM
      // ======================

      if (session.step === 'confirm') {

        if (
          text.toLowerCase() !== 'yes'
        ) {

          return ctx.reply(
            'Type YES to confirm or CANCEL to cancel.'
          );
        }

        const user =
          await getBotUser(ctx.from.id);

        const repName =
          cleanText(
            user.fields.Title ||
            user.fields.LinkTitle
          );

        await acceptTablet(
          session.tabletItemId,
          session.data,
          repName
        );

        delete sessions[userId];

        return ctx.reply(
          '✅ Tablet accepted successfully.'
        );
      }
    }

    // ======================
// CONFIRM EOD FLOW
// ======================

if (session.type === 'confirmEOD') {

  if (session.step === 'confirm') {

    if (text.toLowerCase() === 'yes') {

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

    if (text.toLowerCase() === 'edit') {

      session.type =
        'submitSales';

      session.step =
        'd10';

      return ctx.reply(
        `Enter final $10 donation count. Current: ${session.data.d10}`
      );
    }

    return ctx.reply(
      'Type YES to submit or EDIT to adjust.'
    );
  }
}
    
    // ======================
    // SALES FLOW
    // ======================

    if (session.type === 'submitSales') {

      const validateNumber = value => {

        const num =
          Number(value);

        return (
          Number.isInteger(num) &&
          num >= 0
        )
          ? num
          : null;
      };

      switch (session.step) {

        case 'd10':

          session.data.d10 =
            validateNumber(text);

          if (session.data.d10 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
          }

          session.step = 'd20';

          return ctx.reply(
            'How many $20 donations did you get?'
          );

        case 'd20':

          session.data.d20 =
            validateNumber(text);

          if (session.data.d20 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
          }

          session.step = 'd25';

          return ctx.reply(
            'How many $25 donations did you get?'
          );

        case 'd25':

          session.data.d25 =
            validateNumber(text);

          if (session.data.d25 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
          }

          session.step = 'd30';

          return ctx.reply(
            'How many $30 donations did you get?'
          );

        case 'd30':

          session.data.d30 =
            validateNumber(text);

          if (session.data.d30 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
          }

          session.step = 'd35';

          return ctx.reply(
            'How many $35 donations did you get?'
          );

        case 'd35':

          session.data.d35 =
            validateNumber(text);

          if (session.data.d35 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
          }

          session.step = 'd40';

          return ctx.reply(
            'How many $40 donations did you get?'
          );

        case 'd40':

          session.data.d40 =
            validateNumber(text);

          if (session.data.d40 === null) {

            return ctx.reply(
              'Please enter a valid number.'
            );
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

🔥 Total Donations:
${
  session.data.d10 +
  session.data.d20 +
  session.data.d25 +
  session.data.d30 +
  session.data.d35 +
  session.data.d40
}

Type YES to submit
or CANCEL to cancel.`
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
    }

  } catch (error) {

    console.log(
      error.response?.data ||
      error.message
    );

    return ctx.reply(
      '❌ Action failed. Please contact management.'
    );
  }
});

// ======================
// BOT START
// ======================

bot.catch((err) => {

  console.log(
    'Bot error:',
    err
  );
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
