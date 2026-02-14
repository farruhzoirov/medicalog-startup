import { registerAs } from '@nestjs/config';

export default registerAs('AUTH', () => ({
  LOGIN: process.env.LOGIN,
  PASSWORD: process.env.PASSWORD,
}));
