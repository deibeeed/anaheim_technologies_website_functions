import * as functions from 'firebase-functions';
import axios, {AxiosError} from 'axios';
import {defineSecret} from 'firebase-functions/params';

const microsoftTenantId = defineSecret('MICROSOFT_TENANT_ID');
const microsoftClientId = defineSecret('MICROSOFT_CLIENT_ID');
const microsoftClientSecret = defineSecret('MICROSOFT_CLIENT_SECRET');

// // Start writing functions
// // https://firebase.google.com/docs/functions/typescript
//
export const helloWorld = functions.https.onRequest((request, response) => {
  functions.logger.info('Hello logs!', {structuredData: true});
  response.send('Hello from Firebase!');
});

export const sendMail = functions.runWith({secrets: [
  microsoftTenantId.name,
  microsoftClientId.name,
  microsoftClientSecret.name,
]}).https.onRequest(async (req, res) => {
  // get microsoft access token
  const tenantId = microsoftTenantId.value();
  const clientId = microsoftClientId.value();
  const clientSecret = microsoftClientSecret.value();
  const scope = 'https://graph.microsoft.com/.default';
  try {
    const accessTokenResponse = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      client_id: clientId,
      client_secret: encodeURI(clientSecret),
      scope: encodeURI(scope),
      grant_type: 'client_credentials',
    }, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Host': 'login.microsoftonline.com',
      },
    });

    const accessToken = accessTokenResponse.data.access_token;
    const tokenType = accessTokenResponse.data.token_type;

    const sendMailResponse = await axios.post('https://graph.microsoft.com/v1.0/users/no-reply@anaheimtechnologies.com/sendMail', {
      message: {
        subject: 'Test email from firebase functions',
        toRecipients: [
          {
            emailAddress: {
              address: 'dafduldulao@gmail.com',
            },
          },
        ],
        body: {
          contentType: 'html',
          content: '<b>HELLO!</b> from firebase functions!',
        },
      },
    }, {
      headers: {
        'Authorization': `${tokenType} ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (sendMailResponse.status >= 200 && sendMailResponse.status <= 202) {
      res.send({
        status: 'ok',
      });

      return;
    }

    throw Error(sendMailResponse.statusText);
  } catch (e) {
    if (e instanceof AxiosError) {
      res.send(e.toJSON());
    }

    res.send(e);
  }
});

