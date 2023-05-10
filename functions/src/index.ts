import axios, {AxiosError} from 'axios';
import {defineSecret} from 'firebase-functions/params';
import {onDocumentCreated} from 'firebase-functions/v2/firestore';
import {onRequest} from 'firebase-functions/v2/https';
import * as functions from 'firebase-functions/v2';

const microsoftTenantId = defineSecret('MICROSOFT_TENANT_ID');
const microsoftClientId = defineSecret('MICROSOFT_CLIENT_ID');
const microsoftClientSecret = defineSecret('MICROSOFT_CLIENT_SECRET');

// // Start writing functions
// // https://firebase.google.com/docs/functions/typescript
//
export const sendMail = onRequest({secrets: [
  microsoftTenantId.name,
  microsoftClientId.name,
  microsoftClientSecret.name,
]}, async (req, res) => {
  functions.logger.info('sendMail called');
  // get microsoft access token
  const tenantId = microsoftTenantId.value();
  const clientId = microsoftClientId.value();
  const clientSecret = microsoftClientSecret.value();
  const scope = 'https://graph.microsoft.com/.default';
  // eslint-disable-next-line max-len
  const {isInquiry, subject, messageHtml, emailAddress, messageRaw, fullName} = req.body;

  if (!isInquiry) {
    res
        .status(400)
        .send({
          error: 'Non-inquiry based message not supported yet.',
        });
    return;
  }

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
        subject: subject,
        toRecipients: [
          {
            emailAddress: {
              address: emailAddress,
            },
          },
        ],
        body: {
          contentType: 'html',
          content: messageHtml,
        },
      },
    }, {
      headers: {
        'Authorization': `${tokenType} ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    // send automated message to inquire
    await axios.post('https://graph.microsoft.com/v1.0/users/no-reply@anaheimtechnologies.com/sendMail', {
      message: {
        subject: `[no-reply] An inquiry was made by ${fullName}`,
        toRecipients: [
          {
            emailAddress: {
              address: 'inquire@anaheimtechnologies.com',
            },
          },
        ],
        body: {
          contentType: 'html',
          // eslint-disable-next-line max-len
          content: `An inquiry was made.</br><hr></br>Full name: ${fullName}</br>Email Address: ${emailAddress}</br>Message: ${messageRaw}`,
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
        status: 'success',
      });

      return;
    }

    throw Error(sendMailResponse.statusText);
  } catch (e) {
    if (e instanceof AxiosError) {
      res.status(400).send(e.toJSON());
      return;
    }

    res.status(400).send(e);
  }
});

// eslint-disable-next-line max-len
export const inquiryCreated = onDocumentCreated('inquiries/{docId}', async (event) => {
  functions.logger.info('inquiry created');
  const snapshot = event.data;

  if (!snapshot) {
    // eslint-disable-next-line max-len
    functions.logger.error(`function triggered but no snapshot was retrieved. PATH: ${event.location}`);
    return;
  }

  const {
    email_address: emailAddress,
    full_name: fullName,
    message,
  } = snapshot.data();

  try {
    let url = 'https://sendmail-pk7b6cgija-uc.a.run.app';

    // eslint-disable-next-line max-len
    if (process.env.FIRESTORE_EMULATOR == 'true' || Boolean(process.env.FIRESTORE_EMULATOR)) {
      url = 'http://127.0.0.1:5001/anaheim-technologies/us-central1/sendMail';
    }

    functions.logger.info('url: ' + url);
    const sendMailResponse = await axios.post(url, {
      isInquiry: true,
      subject: 'Thank you for your message!',
      emailAddress: emailAddress,
      messageHtml: composeEmailMessage(fullName),
      messageRaw: message,
      fullName: fullName,
    });

    // eslint-disable-next-line max-len
    functions.logger.info('sendMailResponse: ' + JSON.stringify(sendMailResponse.data));
  } catch (err) {
    functions.logger.error(err);
  }
});

// eslint-disable-next-line require-jsdoc
function composeEmailMessage(fullName: string) {
  // eslint-disable-next-line max-len
  return `Dear ${fullName},</br></br>We received your message!</br>Thank you for taking your time to message us. Rest assured that our team will contact you soon.</br></br></br>Cheers,</br></br><b>Anaheim Technologies, OPC</b>`;
}
