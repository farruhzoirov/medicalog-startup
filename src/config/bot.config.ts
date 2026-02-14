import { registerAs } from '@nestjs/config';

export default registerAs('BOT', () => ({
  TELEGRAM_BOT_TOKEN: process.env.TELEGRAM_BOT_TOKEN,
  ADMIN_1_ID: process.env.ADMIN_1_ID,
  ADMIN_2_ID: process.env.ADMIN_2_ID,
  USER_ID: process.env.USER_ID,
}));
