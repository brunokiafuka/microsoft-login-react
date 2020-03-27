
import { UserAgentApplication } from 'msal';
import { getUserDetails } from './GraphService'

class Microsoft {
    constructor({ appId, redirectUri, scopes }) {
        this.config = {
            appId,
            redirectUri,
            scopes
        }

        this.userAgentApplication = new UserAgentApplication({
            auth: {
                clientId: this.config.appId,
                redirectUri: this.config.redirectUri
            },
            cache: {
                cacheLocation: "localStorage",
                storeAuthStateInCookie: true
            }
        })
    }

    async login() {
        try {
            await userAgentApplication.loginPopup(
                {
                    scopes: MicrosoftConfig.scopes,
                    prompt: "select_account"
                });

            const msUserInfo = await this.getMicrosoftUserProfile()

            return { isSuccess: false, user: msUserInfo }
        } catch (error) {
            return { isSuccess: false, error }
        }
    }

    async getMicrosoftUserProfile() {
        try {
            const accessToken = await userAgentApplication.acquireTokenSilent({
                scopes: this.config.scopes,
            });

            if (accessToken) {
                // Get the user's profile from Graph
                const userData = await getUserDetails(accessToken);
                return userData
            }
        }
        catch (err) {
            let error = {};
            if (typeof (err) === 'string') {
                var errParts = err.split('|');
                error = errParts.length > 1 ?
                    { message: errParts[1], debug: errParts[0] } :
                    { message: err };
            } else {
                error = {
                    message: err.message,
                    debug: JSON.stringify(err)
                };
            }

            return error
        }
    }

}

export default Microsoft()