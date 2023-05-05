import * as functions from 'firebase-functions';
import * as nodemailer from 'nodemailer';

// // Start writing functions
// // https://firebase.google.com/docs/functions/typescript
//
export const helloWorld = functions.https.onRequest((request, response) => {
  functions.logger.info('Hello logs!', {structuredData: true});
  response.send('Hello from Firebase!');
});

export const sendMail = functions.https.onRequest(async (req, res) => {
  const transporter = nodemailer.createTransport();
  const mailResponse = await transporter.sendMail({

  });

  console.log(`Mail send response ${mailResponse.response}`);

});

