import { google } from 'googleapis';
import { clientId, clientSecret, redirectUrl, refreshToken } from '../config';

/**
 * Prepare oauth2 client
 * @type {OAuth2Client}
 */
const auth = new google.auth.OAuth2(clientId, clientSecret, redirectUrl);
auth.setCredentials({
  access_token: 'DUMMY',
  expiry_date: 1,
  refresh_token: refreshToken,
  token_type: 'Bearer',
});

export default auth;
