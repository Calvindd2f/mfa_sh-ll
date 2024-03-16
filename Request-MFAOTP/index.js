const {Twilio} = require('twilio');
const {MSRestAzure, InteractiveLogin} = require('@azure/ms-rest-nodeauth');
const {GraphRbacManagementClient} = require('@azure/graph');
const {Client: GraphClient} = require('@microsoft/microsoft-graph-client');

module.exports = async function (context, req) {
    const userEmail = req.query.email;

    if (!userEmail) {
        context.res = {
            status: 400,
            body: 'Missing email parameter'
        };
        return;
    }

    const clientId = process.env.AZURE_AD_CLIENT_ID;
    const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;
    const tenantId = process.env.AZURE_AD_TENANT_ID;

    try {
        const credentials = await MSRestAzure.loginWithServicePrincipalSecret(clientId, clientSecret, tenantId);
        const graphClient = GraphClient.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => credentials.tokenCache._entries[0].accessToken
            }
        });

        const user = await graphClient.api(`/users/${userEmail}`).get();
        const recipientPhone = user.mobilePhone;

        if (!recipientPhone) {
            context.res = {
                status: 404,
                body: 'Phone number not found for the specified user'
            };
            return;
        }

        const otp = generateOTP();

        const accountSid = process.env.TWILIO_ACCOUNT_SID;
        const authToken = process.env.TWILIO_AUTH_TOKEN;
        const twilioPhone = process.env.TWILIO_PHONE_NUMBER;

        const twilioClient = new Twilio(accountSid, authToken);

        const message = await twilioClient.messages.create({
            body: `Your OTP is: ${otp}`,
            from: twilioPhone,
            to: recipientPhone
        });
        
        
        context.res = {
            status: 200,
            body: {otp}
        };
    } catch (error) {
        context.res = {
            status: 500,
            body: `Error: ${error.message}`
        };
    }
};

function generateOTP() {
    return Math.floor(100000 + Math.random() * 900000);
}
