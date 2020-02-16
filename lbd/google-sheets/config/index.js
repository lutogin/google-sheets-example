const dotenv = require('dotenv-safe');

dotenv.config();

export const clientId = process.env.GOOGLE_SHEETS_CLIENT_ID;
export const clientSecret = process.env.GOOGLE_SHEETS_CLIENT_SECRET;
export const redirectUrl = process.env.GOOGLE_SHEETS_REDIRECT_URL;
export const refreshToken = process.env.GOOGLE_SHEETS_REFRESH_TOKEN;
