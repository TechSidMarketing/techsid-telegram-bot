require('dotenv').config();

const { Telegraf } = require('telegraf');

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

bot.launch();

console.log('TechSid Telegram Bot is running...');
