import TelegramBot from "node-telegram-bot-api";

export const sendMessage = async (messageText: string, telegramBotToken: string, admin1Id: string, admin2Id: string) => {
  const bot = new TelegramBot(telegramBotToken);
  if (admin1Id) {
    await bot.sendMessage(admin1Id, messageText, { parse_mode: "HTML" });
  }
  if (admin2Id) {
    await bot.sendMessage(admin2Id, messageText, { parse_mode: "HTML" });
  }
};