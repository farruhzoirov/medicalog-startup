import { registerAs } from '@nestjs/config';

export default registerAs('CONFIG_DATABASE', () => ({
  MONGODB_URI: process.env.MONGODB_URI,
}));
