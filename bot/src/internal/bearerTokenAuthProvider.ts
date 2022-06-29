import { AuthenticationProvider, AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';

export class BearerTokenAuthProvider implements AuthenticationProvider {
    private _accessToken: string;

    constructor(accessToken: string) {
        this._accessToken = accessToken;
    }

    public async getAccessToken(authenticationProviderOptions?: AuthenticationProviderOptions): Promise<string> {
        return this._accessToken;
    }
}